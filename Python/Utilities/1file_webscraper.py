#!/usr/bin/env python3
"""
Advanced web scraper with caching, checkpoints, and game title extraction.
"""

import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import csv
import random
import argparse
import logging
from urllib.robotparser import RobotFileParser
from urllib.parse import urljoin, urlparse
import re
from tqdm import tqdm
import sys
import json
import os
import hashlib
from pathlib import Path
import pickle

class CacheManager:
    """Manages local HTML caching to reduce network requests."""
    
    def __init__(self, cache_dir="scraper_cache", max_age_hours=24):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(exist_ok=True)
        self.max_age = max_age_hours * 3600  # Convert to seconds
        
    def _get_cache_path(self, url):
        """Generate cache file path for a URL."""
        url_hash = hashlib.md5(url.encode()).hexdigest()
        return self.cache_dir / f"{url_hash}.html"
    
    def get(self, url):
        """Retrieve cached HTML if available and not expired."""
        cache_path = self._get_cache_path(url)
        
        if not cache_path.exists():
            return None
            
        # Check if cache is expired
        if time.time() - cache_path.stat().st_mtime > self.max_age:
            cache_path.unlink()  # Remove expired cache
            return None
            
        try:
            return cache_path.read_text(encoding='utf-8')
        except Exception:
            return None
    
    def set(self, url, html):
        """Store HTML in cache."""
        try:
            cache_path = self._get_cache_path(url)
            cache_path.write_text(html, encoding='utf-8')
        except Exception as e:
            logging.warning(f"Failed to cache {url}: {e}")
    
    def clear(self):
        """Clear all cached files."""
        for cache_file in self.cache_dir.glob("*.html"):
            cache_file.unlink()
        logging.info("Cache cleared")
    
    def get_stats(self):
        """Get cache statistics."""
        files = list(self.cache_dir.glob("*.html"))
        total_size = sum(f.stat().st_size for f in files)
        return {
            'files': len(files),
            'size_mb': total_size / (1024 * 1024),
            'cache_dir': str(self.cache_dir)
        }


class CheckpointManager:
    """Manages scraping checkpoints for resumable operations."""
    
    def __init__(self, checkpoint_file="scraper_checkpoint.json"):
        self.checkpoint_file = Path(checkpoint_file)
        self.data = self._load_checkpoint()
    
    def _load_checkpoint(self):
        """Load existing checkpoint data."""
        if self.checkpoint_file.exists():
            try:
                with open(self.checkpoint_file, 'r') as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            'completed_pages': [],
            'game_links': [],
            'fichier_links': [],
            'last_page': 0,
            'timestamp': None
        }
    
    def save_checkpoint(self, completed_pages, game_links, fichier_links, last_page):
        """Save current progress."""
        self.data.update({
            'completed_pages': completed_pages,
            'game_links': game_links,
            'fichier_links': fichier_links,
            'last_page': last_page,
            'timestamp': time.time()
        })
        
        try:
            with open(self.checkpoint_file, 'w') as f:
                json.dump(self.data, f, indent=2)
        except Exception as e:
            logging.warning(f"Failed to save checkpoint: {e}")
    
    def get_resume_data(self):
        """Get data for resuming scraping."""
        return self.data
    
    def clear_checkpoint(self):
        """Remove checkpoint file."""
        if self.checkpoint_file.exists():
            self.checkpoint_file.unlink()
        self.data = {
            'completed_pages': [],
            'game_links': [],
            'fichier_links': [],
            'last_page': 0,
            'timestamp': None
        }


class GameScraper:
    def __init__(self, config):
        self.config = config
        self.session = requests.Session()
        self.session.headers.update(self.get_realistic_headers())
        self.game_page_links = []
        self.fichier_links = []
        self.setup_logging()
        
        # Initialize managers
        self.cache = CacheManager(config.cache_dir, config.cache_hours)
        self.checkpoint = CheckpointManager(config.checkpoint_file)
        
        # Track completed pages for checkpointing
        self.completed_pages = []
        
    def get_realistic_headers(self):
        """Return realistic browser headers to avoid detection."""
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Cache-Control": "max-age=0"
        }
        return headers
    
    def setup_logging(self):
        """Configure logging with appropriate level and format."""
        log_level = getattr(logging, self.config.log_level.upper())
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('scraper.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def check_robots_txt(self, base_url):
        """Check if scraping is allowed according to robots.txt."""
        try:
            robots_url = urljoin(base_url, '/robots.txt')
            rp = RobotFileParser()
            rp.set_url(robots_url)
            rp.read()
            
            user_agent = self.session.headers.get('User-Agent', '*')
            can_fetch = rp.can_fetch(user_agent, base_url)
            
            if not can_fetch:
                self.logger.warning(f"robots.txt disallows scraping for {base_url}")
                if not self.config.ignore_robots:
                    self.logger.error("Scraping blocked by robots.txt. Use --ignore-robots to override.")
                    return False
            else:
                self.logger.info("robots.txt allows scraping")
            
            # Get crawl delay if specified
            crawl_delay = rp.crawl_delay(user_agent)
            if crawl_delay:
                self.config.min_delay = max(self.config.min_delay, crawl_delay)
                self.logger.info(f"robots.txt specifies crawl delay: {crawl_delay}s")
                
        except Exception as e:
            self.logger.warning(f"Could not check robots.txt: {e}")
            
        return True
    
    def detect_max_pages(self):
        """Dynamically detect the maximum number of pages."""
        try:
            # Try to find pagination info on first page
            html = self.fetch_page(self.config.base_url.format(1))
            if not html:
                return self.config.max_pages
                
            soup = BeautifulSoup(html, 'html.parser')
            
            # Common pagination selectors
            pagination_selectors = [
                '.pagination a',
                '.page-numbers a',
                '.nav-links a',
                'a[href*="page/"]'
            ]
            
            max_page = 1
            for selector in pagination_selectors:
                links = soup.select(selector)
                for link in links:
                    href = link.get('href', '')
                    # Extract page number from URL
                    page_match = re.search(r'/page/(\d+)/', href)
                    if page_match:
                        page_num = int(page_match.group(1))
                        max_page = max(max_page, page_num)
            
            if max_page > 1:
                self.logger.info(f"Detected maximum pages: {max_page}")
                return min(max_page, self.config.max_pages)  # Cap at config limit
            
        except Exception as e:
            self.logger.warning(f"Could not detect max pages: {e}")
            
        return self.config.max_pages
    
    def fetch_page(self, url, retries=None, use_cache=True):
        """Request a webpage with caching, retry and exponential backoff."""
        retries = retries or self.config.retries
        
        # Try cache first
        if use_cache:
            cached_html = self.cache.get(url)
            if cached_html:
                self.logger.debug(f"Cache hit for {url}")
                return cached_html
        
        # Fetch from network
        for attempt in range(retries):
            try:
                response = self.session.get(url, timeout=self.config.timeout)
                response.raise_for_status()
                html = response.text
                
                # Cache the result
                if use_cache:
                    self.cache.set(url, html)
                
                return html
                
            except requests.exceptions.RequestException as e:
                self.logger.warning(f"Attempt {attempt + 1} failed for {url}: {e}")
                if attempt < retries - 1:
                    delay = self.config.retry_delay * (2 ** attempt)
                    time.sleep(delay)
                else:
                    self.logger.error(f"Failed to fetch {url} after {retries} attempts")
                    
        return None
    
    def parse_game_links_with_titles(self, html):
        """Extract game post links AND titles from a category page."""
        soup = BeautifulSoup(html, 'html.parser')
        games = []
        
        for post_element in soup.select(self.config.post_selector):
            link_element = post_element.find('a')
            if link_element and link_element.get('href'):
                href = link_element.get('href')
                title = link_element.get_text(strip=True)
                
                if href and self.validate_url(href):
                    games.append({
                        'url': href,
                        'title': title or 'Unknown Title'
                    })
                    
        return games
    
    def validate_url(self, url):
        """Validate that URL is properly formatted and from expected domain."""
        try:
            parsed = urlparse(url)
            if not parsed.scheme or not parsed.netloc:
                return False
            
            # Check if URL is from expected domain (optional)
            base_domain = urlparse(self.config.base_url).netloc
            if self.config.strict_domain and parsed.netloc != base_domain:
                return False
                
            return True
        except Exception:
            return False
    
    def validate_fichier_link(self, url):
        """Validate 1fichier links more strictly."""
        if not url or not isinstance(url, str):
            return False
            
        url = url.strip()
        
        # Check if it's a valid 1fichier URL
        if not re.match(r'https?://1fichier\.com/\?[a-zA-Z0-9]+', url):
            return False
            
        return self.validate_url(url)
    
    def categorize_link(self, link_element, context_text=""):
        """Analyze link context to determine type, version, and quality."""
        # Get text from the link and surrounding context
        link_text = link_element.get_text(strip=True).lower()
        
        # Look for context in parent elements and siblings
        parent = link_element.parent
        context_sources = [link_text, context_text.lower()]
        
        # Add text from parent and sibling elements for context
        if parent:
            context_sources.append(parent.get_text(strip=True).lower())
            
        # Check previous and next siblings for context
        prev_sibling = link_element.previous_sibling
        next_sibling = link_element.next_sibling
        
        if prev_sibling and hasattr(prev_sibling, 'get_text'):
            context_sources.append(prev_sibling.get_text(strip=True).lower())
        if next_sibling and hasattr(next_sibling, 'get_text'):
            context_sources.append(next_sibling.get_text(strip=True).lower())
            
        # Combine all context
        full_context = " ".join(context_sources)
        
        # Initialize result
        result = {
            'link_type': 'Base Game',  # Default
            'version': '',
            'quality': '',
            'platform': '',
            'region': '',
            'language': ''
        }
        
        # Detect link type
        if any(keyword in full_context for keyword in ['update', 'patch', 'upd']):
            result['link_type'] = 'Update'
        elif any(keyword in full_context for keyword in ['dlc', 'downloadable content', 'expansion']):
            result['link_type'] = 'DLC'
        elif any(keyword in full_context for keyword in ['addon', 'add-on']):
            result['link_type'] = 'Add-on'
        elif 'demo' in full_context:
            result['link_type'] = 'Demo'
        elif any(keyword in full_context for keyword in ['theme', 'avatar']):
            result['link_type'] = 'Theme/Avatar'
        
        # Extract version numbers
        version_patterns = [
            r'(?:update|patch|ver|version|v)?\s*(\d+\.\d+(?:\.\d+)?)',
            r'(\d+\.\d+(?:\.\d+)?)\s*(?:update|patch)',
            r'v(\d+\.\d+(?:\.\d+)?)',
        ]
        
        for pattern in version_patterns:
            version_match = re.search(pattern, full_context, re.IGNORECASE)
            if version_match:
                result['version'] = version_match.group(1)
                break
        
        # Extract quality/resolution
        quality_patterns = [
            r'(\d{3,4}p)',  # 720p, 1080p, etc.
            r'(\d{3,4}x\d{3,4})',  # 1920x1080
            r'(4k|uhd|hd)',
        ]
        
        for pattern in quality_patterns:
            quality_match = re.search(pattern, full_context, re.IGNORECASE)
            if quality_match:
                result['quality'] = quality_match.group(1).upper()
                break
        
        # Extract platform info
        platforms = ['ps4', 'ps5', 'xbox', 'pc', 'switch', 'vita']
        for platform in platforms:
            if platform in full_context:
                result['platform'] = platform.upper()
                break
        
        # Extract region
        regions = ['eur', 'usa', 'jpn', 'asia', 'pal', 'ntsc']
        for region in regions:
            if region in full_context and len(region) >= 3:  # Avoid false matches
                result['region'] = region.upper()
                break
        
        # Extract file host preference (Mediafire, Viki, Akia, etc.)
        hosts = ['mediafire', 'viki', 'akia', 'mega', 'google drive', 'dropbox']
        for host in hosts:
            if host in full_context:
                result['quality'] = result['quality'] or host.title()
                break
        
        return result

    def extract_1fichier_links(self, game_data):
        """Scrape and categorize 1fichier.com links from a game post."""
        game_url = game_data['url']
        game_title = game_data['title']
        
        html = self.fetch_page(game_url)
        if not html:
            return []
        
        soup = BeautifulSoup(html, 'html.parser')
        content = soup.find("div", class_="entry")
        if not content:
            return []
        
        valid_links = []
        
        # Get all text content for context analysis
        full_text = content.get_text()
        
        # Find all paragraphs or line breaks to understand structure
        text_blocks = []
        for element in content.find_all(['p', 'br', 'div', 'span']):
            if element.get_text(strip=True):
                text_blocks.append(element.get_text(strip=True))
        
        # Process each link with its context
        for a in content.find_all("a", href=True):
            href = a['href'].strip()
            
            if "1fichier.com" in href.lower() and self.validate_fichier_link(href):
                # Find the text block containing this link for context
                link_context = ""
                for block in text_blocks:
                    if a.get_text(strip=True) in block:
                        link_context = block
                        break
                
                # Categorize the link
                link_info = self.categorize_link(a, link_context)
                
                # Build the result
                link_data = {
                    "game_title": game_title,
                    "link_type": link_info['link_type'],
                    "version": link_info['version'],
                    "quality": link_info['quality'],
                    "platform": link_info['platform'],
                    "region": link_info['region'],
                    "game_page": game_url,
                    "1fichier_link": href
                }
                
                # Only add if it matches our filtering criteria
                if self.should_include_link(link_data):
                    valid_links.append(link_data)
        
        return valid_links
    
    def load_resume_data(self):
        """Load checkpoint data for resuming."""
        resume_data = self.checkpoint.get_resume_data()
        
        if resume_data['timestamp'] and resume_data['last_page'] > 0:
            age_hours = (time.time() - resume_data['timestamp']) / 3600
            self.logger.info(f"Found checkpoint from {age_hours:.1f} hours ago (page {resume_data['last_page']})")
            
            if not self.config.resume or age_hours > 48:  # Auto-expire old checkpoints
                self.logger.info("Checkpoint too old or resume disabled, starting fresh")
                return None
            
            # Convert old format game_links (URLs) to new format (dicts)
            game_links = []
            for item in resume_data.get('game_links', []):
                if isinstance(item, str):
                    # Old format - just URL
                    game_links.append({'url': item, 'title': 'Unknown Title'})
                elif isinstance(item, dict):
                    # New format - already has title
                    game_links.append(item)
            
            self.game_page_links = game_links
            self.fichier_links = resume_data.get('fichier_links', [])
            self.completed_pages = resume_data.get('completed_pages', [])
            
            self.logger.info(f"Resuming with {len(self.game_page_links)} game links, {len(self.fichier_links)} download links")
            return resume_data
        
        return None
    
    def collect_game_post_urls(self):
        """Scrape all category pages for game post URLs with checkpoint support."""
        max_pages = self.detect_max_pages() if self.config.auto_detect_pages else self.config.max_pages
        
        # Try to resume from checkpoint
        resume_data = self.load_resume_data()
        start_page = resume_data['last_page'] + 1 if resume_data else 1
        
        if start_page > max_pages:
            self.logger.info("All pages already processed according to checkpoint")
            return
        
        self.logger.info(f"Collecting game links from pages {start_page}-{max_pages}...")
        
        with tqdm(total=max_pages - start_page + 1, desc="Collecting game links") as pbar:
            for page in range(start_page, max_pages + 1):
                url = self.config.base_url.format(page)
                
                html = self.fetch_page(url)
                if not html:
                    self.logger.warning(f"Failed to fetch page {page}")
                    pbar.update(1)
                    continue
                
                games = self.parse_game_links_with_titles(html)
                self.game_page_links.extend(games)
                self.completed_pages.append(page)
                
                pbar.set_postfix({
                    'Total Games': len(self.game_page_links),
                    'Page Games': len(games),
                    'Cache': f"{self.cache.get_stats()['files']} files"
                })
                pbar.update(1)
                
                # Save checkpoint every 10 pages
                if page % 10 == 0:
                    self.checkpoint.save_checkpoint(
                        self.completed_pages,
                        self.game_page_links,
                        self.fichier_links,
                        page
                    )
                
                # Respect rate limiting
                delay = random.uniform(self.config.min_delay, self.config.max_delay)
                time.sleep(delay)
        
        # Final checkpoint save
        self.checkpoint.save_checkpoint(
            self.completed_pages,
            self.game_page_links,
            self.fichier_links,
            max_pages
        )
    
    def collect_all_fichier_links(self):
        """Use threads to speed up extraction of 1fichier links."""
        if not self.game_page_links:
            self.logger.warning("No game links found to process")
            return
        
        # Filter out games we've already processed (for resume functionality)
        processed_urls = {link.get('game_page', '') for link in self.fichier_links}
        remaining_games = [
            game for game in self.game_page_links 
            if game['url'] not in processed_urls
        ]
        
        if not remaining_games:
            self.logger.info("All games already processed for download links")
            return
            
        self.logger.info(f"Extracting 1fichier links from {len(remaining_games)} remaining games...")
        
        with ThreadPoolExecutor(max_workers=self.config.max_workers) as executor:
            futures = {
                executor.submit(self.extract_1fichier_links, game): game 
                for game in remaining_games
            }
            
            with tqdm(total=len(futures), desc="Extracting download links") as pbar:
                processed_count = 0
                for future in as_completed(futures):
                    try:
                        links = future.result()
                        if links:
                            self.fichier_links.extend(links)
                    except Exception as e:
                        game = futures[future]
                        self.logger.error(f"Error processing {game['title']}: {e}")
                    
                    processed_count += 1
                    pbar.set_postfix({
                        'Total Links': len(self.fichier_links),
                        'Cache': f"{self.cache.get_stats()['files']} files"
                    })
                    pbar.update(1)
                    
                    # Save checkpoint every 50 processed games
                    if processed_count % 50 == 0:
                        self.checkpoint.save_checkpoint(
                            self.completed_pages,
                            self.game_page_links,
                            self.fichier_links,
                            max(self.completed_pages) if self.completed_pages else 0
                        )
    
    def should_include_link(self, link_data):
        """Filter links based on configuration preferences."""
        # Apply filters based on config
        if hasattr(self.config, 'link_types') and self.config.link_types:
            if link_data['link_type'].lower() not in [t.lower() for t in self.config.link_types]:
                return False
        
        if hasattr(self.config, 'exclude_updates') and self.config.exclude_updates:
            if link_data['link_type'] == 'Update':
                return False
                
        if hasattr(self.config, 'exclude_dlc') and self.config.exclude_dlc:
            if link_data['link_type'] == 'DLC':
                return False
        
        if hasattr(self.config, 'min_quality') and self.config.min_quality:
            quality = link_data['quality'].lower()
            if '1080p' in quality or 'uhd' in quality or '4k' in quality:
                quality_score = 3
            elif '720p' in quality:
                quality_score = 2
            elif '480p' in quality:
                quality_score = 1
            else:
                quality_score = 0
                
            min_score = {'480p': 1, '720p': 2, '1080p': 3}.get(self.config.min_quality, 0)
            if quality_score < min_score:
                return False
        
        return True

    def save_to_csv(self):
        """Save collected links to CSV file with enhanced categorization."""
        if not self.fichier_links:
            self.logger.warning("No links to save")
            return
            
        try:
            fieldnames = [
                "game_title", "link_type", "version", "quality", 
                "platform", "region", "game_page", "1fichier_link"
            ]
            
            with open(self.config.output_file, "w", newline='', encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.fichier_links)
            
            self.logger.info(f"Saved {len(self.fichier_links)} categorized links to {self.config.output_file}")
            
            # Print categorization summary
            self.print_categorization_stats()
            
        except Exception as e:
            self.logger.error(f"Failed to save CSV: {e}")
    
    def print_categorization_stats(self):
        """Print statistics about link categorization."""
        if not self.fichier_links:
            return
            
        # Count by link type
        type_counts = {}
        version_counts = {}
        quality_counts = {}
        
        for link in self.fichier_links:
            # Count link types
            link_type = link.get('link_type', 'Unknown')
            type_counts[link_type] = type_counts.get(link_type, 0) + 1
            
            # Count versions
            version = link.get('version', '')
            if version:
                version_counts[version] = version_counts.get(version, 0) + 1
            
            # Count qualities
            quality = link.get('quality', '')
            if quality:
                quality_counts[quality] = quality_counts.get(quality, 0) + 1
        
        print(f"\n📊 Link Categorization Summary:")
        print(f"   Total links: {len(self.fichier_links)}")
        
        print(f"\n🎮 By Link Type:")
        for link_type, count in sorted(type_counts.items()):
            print(f"   {link_type}: {count}")
        
        if version_counts:
            print(f"\n🔄 By Version:")
            for version, count in sorted(version_counts.items()):
                print(f"   v{version}: {count}")
        
        if quality_counts:
            print(f"\n📺 By Quality/Host:")
            for quality, count in sorted(quality_counts.items()):
                print(f"   {quality}: {count}")
    
    def print_stats(self):
        """Print scraping and cache statistics."""
        cache_stats = self.cache.get_stats()
        
        print(f"\n📊 Scraping Statistics:")
        print(f"   Games found: {len(self.game_page_links)}")
        print(f"   Download links: {len(self.fichier_links)}")
        print(f"   Pages processed: {len(self.completed_pages)}")
        print(f"\n💾 Cache Statistics:")
        print(f"   Cached files: {cache_stats['files']}")
        print(f"   Cache size: {cache_stats['size_mb']:.1f} MB")
        print(f"   Cache location: {cache_stats['cache_dir']}")
    
    def run(self):
        """Main execution method."""
        start_time = time.time()
        
        self.logger.info("Starting advanced scraper...")
        
        # Print cache stats if resuming
        if self.config.resume:
            cache_stats = self.cache.get_stats()
            if cache_stats['files'] > 0:
                self.logger.info(f"Using cache with {cache_stats['files']} files ({cache_stats['size_mb']:.1f} MB)")
        
        # Check robots.txt compliance
        base_url = self.config.base_url.format(1)
        if not self.check_robots_txt(base_url):
            return
        
        try:
            # Collect game post URLs
            self.collect_game_post_urls()
            self.logger.info(f"Found {len(self.game_page_links)} total games")
            
            # Extract 1fichier links
            self.collect_all_fichier_links()
            self.logger.info(f"Extracted {len(self.fichier_links)} download links")
            
            # Save results
            self.save_to_csv()
            
            # Clean up checkpoint on successful completion
            if not self.config.keep_checkpoint:
                self.checkpoint.clear_checkpoint()
            
            # Print final statistics
            self.print_stats()
            
        except KeyboardInterrupt:
            self.logger.info("Scraping interrupted - checkpoint saved")
            # Save final checkpoint
            self.checkpoint.save_checkpoint(
                self.completed_pages,
                self.game_page_links,
                self.fichier_links,
                max(self.completed_pages) if self.completed_pages else 0
            )
            raise
        
        elapsed = time.time() - start_time
        self.logger.info(f"Scraping completed in {elapsed:.2f} seconds")


class Config:
    def __init__(self, args):
        self.base_url = args.base_url
        self.post_selector = args.post_selector
        self.max_pages = args.max_pages
        self.max_workers = args.max_workers
        self.timeout = args.timeout
        self.retries = args.retries
        self.retry_delay = args.retry_delay
        self.min_delay = args.min_delay
        self.max_delay = args.max_delay
        self.output_file = args.output_file
        self.log_level = args.log_level
        self.ignore_robots = args.ignore_robots
        self.auto_detect_pages = args.auto_detect_pages
        self.strict_domain = args.strict_domain
        
        # New link categorization and filtering options
        self.link_types = getattr(args, 'link_types', None)
        self.exclude_updates = getattr(args, 'exclude_updates', False)
        self.exclude_dlc = getattr(args, 'exclude_dlc', False)
        self.min_quality = getattr(args, 'min_quality', None)
        self.cache_dir = args.cache_dir
        self.cache_hours = args.cache_hours
        self.checkpoint_file = args.checkpoint_file
        self.resume = args.resume
        self.keep_checkpoint = args.keep_checkpoint
        self.clear_cache = args.clear_cache


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Advanced web scraper with caching, checkpoints, and game title extraction",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    # Original arguments
    parser.add_argument(
        "--base-url",
        default="https://dlpsgame.com/category/ps4/page/{}/",
        help="Base URL pattern with {} placeholder for page number"
    )
    
    parser.add_argument(
        "--post-selector",
        default="h2.post-title",
        help="CSS selector for post titles"
    )
    
    parser.add_argument(
        "--max-pages",
        type=int,
        default=296,
        help="Maximum number of pages to scrape"
    )
    
    parser.add_argument(
        "--max-workers",
        type=int,
        default=10,
        help="Maximum number of concurrent threads"
    )
    
    parser.add_argument(
        "--timeout",
        type=int,
        default=15,
        help="Request timeout in seconds"
    )
    
    parser.add_argument(
        "--retries",
        type=int,
        default=3,
        help="Number of retry attempts for failed requests"
    )
    
    parser.add_argument(
        "--retry-delay",
        type=float,
        default=2.0,
        help="Base delay between retries (exponential backoff)"
    )
    
    parser.add_argument(
        "--min-delay",
        type=float,
        default=0.5,
        help="Minimum delay between requests"
    )
    
    parser.add_argument(
        "--max-delay",
        type=float,
        default=1.0,
        help="Maximum delay between requests"
    )
    
    parser.add_argument(
        "--output-file",
        default="ps4_1fichier_links.csv",
        help="Output CSV filename"
    )
    
    parser.add_argument(
        "--log-level",
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help="Logging level"
    )
    
    parser.add_argument(
        "--ignore-robots",
        action='store_true',
        help="Ignore robots.txt restrictions"
    )
    
    parser.add_argument(
        "--auto-detect-pages",
        action='store_true',
        help="Automatically detect maximum pages"
    )
    
    parser.add_argument(
        "--strict-domain",
        action='store_true',
        help="Only allow URLs from the base domain"
    )
    
    # New caching and checkpoint arguments
    parser.add_argument(
        "--cache-dir",
        default="scraper_cache",
        help="Directory for HTML cache files"
    )
    
    parser.add_argument(
        "--cache-hours",
        type=int,
        default=24,
        help="Hours to keep cached HTML files"
    )
    
    parser.add_argument(
        "--checkpoint-file",
        default="scraper_checkpoint.json",
        help="Checkpoint file for resumable scraping"
    )
    
    parser.add_argument(
        "--resume",
        action='store_true',
        default=True,
        help="Resume from checkpoint if available"
    )
    
    parser.add_argument(
        "--no-resume",
        action='store_false',
        dest='resume',
        help="Start fresh, ignore checkpoints"
    )
    
    parser.add_argument(
        "--keep-checkpoint",
        action='store_true',
        help="Keep checkpoint file after successful completion"
    )
    
    parser.add_argument(
        "--clear-cache",
        action='store_true',
        help="Clear HTML cache before starting"
    )
    
    # Link categorization and filtering arguments
    parser.add_argument(
        "--link-types",
        nargs='+',
        choices=['base game', 'update', 'dlc', 'add-on', 'demo', 'theme/avatar'],
        help="Only include specific link types (default: all types)"
    )
    
    parser.add_argument(
        "--exclude-updates",
        action='store_true',
        help="Exclude update/patch links"
    )
    
    parser.add_argument(
        "--exclude-dlc",
        action='store_true',
        help="Exclude DLC links"
    )
    
    parser.add_argument(
        "--min-quality",
        choices=['480p', '720p', '1080p'],
        help="Minimum quality threshold for links"
    )
    
    return parser.parse_args()


if __name__ == "__main__":
    try:
        args = parse_arguments()
        config = Config(args)
        scraper = GameScraper(config)
        
        # Clear cache if requested
        if config.clear_cache:
            scraper.cache.clear()
        
        scraper.run()
        
    except KeyboardInterrupt:
        print("\nScraping interrupted by user")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)