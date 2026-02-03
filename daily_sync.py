#!/usr/bin/env python3
"""
ParentMail Daily Calendar Sync
Automatically scrapes ParentMail newsletters and syncs school events to Google Calendar.

This script is designed to run via GitHub Actions on a daily schedule.
Credentials are passed via environment variables for security.

Children:
- Rivan: Year 2 (Y2), Yellow class
- Arvi: Reception (YR), Red class
- Both: KS1 events apply to both children
"""

import os
import re
import json
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple

from playwright.sync_api import sync_playwright, Page, Browser
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Constants
PARENTMAIL_URL = "https://pmx.parentmail.co.uk"
PARENTMAIL_EMAIL = os.environ.get('PARENTMAIL_EMAIL', 'sachinsharma0787@gmail.com')
PARENTMAIL_PASSWORD = os.environ.get('PARENTMAIL_PASSWORD')

# Google Calendar color IDs
COLOR_ARVI = '6'    # Orange for YR/Arvi events
COLOR_RIVAN = '9'   # Blue for Y2/Rivan events
COLOR_BOTH = '11'   # Red for KS1/Both events

# Event filtering keywords
INCLUDE_KEYWORDS = ['yr', 'y2', 'ks1', 'red class', 'yellow class', 'reception', 'year 2']
EXCLUDE_ONLY_KEYWORDS = ['ks2', 'y3', 'y4', 'y5', 'y6', 'year 3', 'year 4', 'year 5', 'year 6']


class ParentMailScraper:
    """Handles ParentMail login and newsletter scraping."""

    def __init__(self, email: str, password: str):
        self.email = email
        self.password = password
        self.browser: Optional[Browser] = None
        self.page: Optional[Page] = None

    def __enter__(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=True)
        self.page = self.browser.new_page()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def login(self) -> bool:
        """Log into ParentMail via IRIS OAuth flow."""
        logger.info("Logging into ParentMail...")

        try:
            self.page.goto(PARENTMAIL_URL)
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(2000)
            self.page.screenshot(path='step1_parentmail_home.png')
            logger.info(f"Step 1 - ParentMail home, URL: {self.page.url}")

            # Click login/sign in if needed
            try:
                login_button = self.page.locator('text=Sign in').first
                if login_button.is_visible(timeout=5000):
                    login_button.click()
                    self.page.wait_for_load_state('networkidle')
                    self.page.wait_for_timeout(2000)
            except:
                pass

            self.page.screenshot(path='step2_after_signin_click.png')
            logger.info(f"Step 2 - After sign in click, URL: {self.page.url}")

            # Step 1: Fill email on ParentMail or IRIS page
            email_field = self.page.locator('input[type="email"], input[name="username"], input[name="email"], #email, #okta-signin-username').first
            email_field.wait_for(state='visible', timeout=10000)
            email_field.fill(self.email)
            logger.info("Filled email field")

            self.page.screenshot(path='step3_email_filled.png')

            # Click continue/next/sign in to proceed to password
            try:
                next_btn = self.page.locator('input[type="submit"], button[type="submit"], button:has-text("Next"), button:has-text("Continue"), button:has-text("Sign in")').first
                next_btn.click()
                logger.info("Clicked next/continue button")
            except Exception as e:
                logger.warning(f"No next button found, trying to continue: {e}")

            # Wait for redirect to IRIS identity provider
            self.page.wait_for_timeout(5000)
            self.page.wait_for_load_state('networkidle')
            self.page.screenshot(path='step4_after_email_submit.png')
            logger.info(f"Step 4 - After email submit, URL: {self.page.url}")

            # Log page content for debugging
            page_text = self.page.locator('body').inner_text()
            logger.info(f"Page contains text (first 500 chars): {page_text[:500]}")

            # IRIS has a two-step login: email first, then password
            # Check if we're on the IRIS email step (shows "Email address" and "Next")
            if 'identity.iris.co.uk' in self.page.url:
                logger.info("On IRIS identity page - handling two-step login")

                # Step 2a: Fill email on IRIS page and click Next
                try:
                    iris_email_field = self.page.locator('input[type="email"], input[type="text"]').first
                    if iris_email_field.is_visible(timeout=3000):
                        # Clear and fill email
                        iris_email_field.fill(self.email)
                        logger.info("Filled email on IRIS page")
                        self.page.screenshot(path='step5_iris_email_filled.png')

                        # Click Next button
                        next_btn = self.page.locator('button:has-text("Next"), input[value="Next"], button[type="submit"]').first
                        next_btn.click()
                        logger.info("Clicked Next on IRIS page")

                        # Wait for password page to load
                        self.page.wait_for_timeout(3000)
                        self.page.wait_for_load_state('networkidle')
                        self.page.screenshot(path='step5b_after_iris_next.png')
                        logger.info(f"After IRIS Next, URL: {self.page.url}")
                except Exception as e:
                    logger.warning(f"IRIS email step handling: {e}")

            # Step 2b: Handle password page
            password_selectors = [
                'input[type="password"]',
                '#okta-signin-password',
                '#password',
                'input[name="password"]',
                'input[name="credentials.passcode"]',
                '[data-se="o-form-input-password"]',
            ]

            password_field = None
            for selector in password_selectors:
                try:
                    field = self.page.locator(selector).first
                    if field.is_visible(timeout=3000):
                        password_field = field
                        logger.info(f"Found password field with selector: {selector}")
                        break
                except:
                    continue

            if not password_field:
                logger.info("Password field not immediately visible, checking page state...")
                self.page.screenshot(path='step5c_looking_for_password.png')

                # Log what's on the page now
                page_text2 = self.page.locator('body').inner_text()
                logger.info(f"Current page text (first 500 chars): {page_text2[:500]}")

                # Try waiting a bit longer and check again
                self.page.wait_for_timeout(3000)

                for selector in password_selectors:
                    try:
                        field = self.page.locator(selector).first
                        if field.is_visible(timeout=3000):
                            password_field = field
                            logger.info(f"Found password field after waiting: {selector}")
                            break
                    except:
                        continue

            if not password_field:
                logger.error("Could not find password field with any selector")
                self.page.screenshot(path='step5_password_not_found.png')
                html_content = self.page.content()
                logger.info(f"Page HTML (first 2000 chars): {html_content[:2000]}")
                return False

            password_field.fill(self.password)
            logger.info("Filled password field")
            self.page.screenshot(path='step6_password_filled.png')

            # Click sign in / verify button
            submit_selectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                'input[value="Sign in"]',
                'input[value="Verify"]',
                'button:has-text("Sign in")',
                'button:has-text("Verify")',
                '[data-se="o-form-button-bar"] button',
            ]

            for selector in submit_selectors:
                try:
                    submit_btn = self.page.locator(selector).first
                    if submit_btn.is_visible(timeout=2000):
                        submit_btn.click()
                        logger.info(f"Clicked submit with selector: {selector}")
                        break
                except:
                    continue

            # Wait for redirect back to ParentMail
            self.page.wait_for_timeout(5000)
            self.page.wait_for_load_state('networkidle')
            self.page.screenshot(path='step7_after_login_submit.png')
            logger.info(f"Step 7 - After login submit, URL: {self.page.url}")

            # Check if login was successful
            if 'pmx.parentmail.co.uk' in self.page.url:
                logger.info("Login successful - redirected to ParentMail!")
                return True

            # Alternative check: look for dashboard elements
            try:
                if self.page.locator('text=Emails').is_visible(timeout=5000):
                    logger.info("Login successful - found Emails link!")
                    return True
            except:
                pass

            # Check if we're stuck on login page
            if 'identity.iris.co.uk' in self.page.url or 'login' in self.page.url.lower():
                logger.error("Login may have failed - still on login/identity page")
                self.page.screenshot(path='login_failed.png')
                return False

            logger.info("Login appears successful")
            return True

        except Exception as e:
            logger.error(f"Login failed: {e}")
            try:
                self.page.screenshot(path='login_error.png')
            except:
                pass
            return False

    def get_latest_newsletter(self) -> Optional[str]:
        """Navigate to emails and find the latest parent newsletter."""
        logger.info("Finding latest newsletter...")

        try:
            # Navigate to emails section
            self.page.goto(f"{PARENTMAIL_URL}/ui/#/messages/emails")
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(2000)

            # Look for newsletter with "Parent newsletter" in title
            newsletter_selector = 'text=/Parent newsletter.*Friday/i'

            # Try to find and click the latest newsletter
            newsletters = self.page.locator(newsletter_selector)
            if newsletters.count() > 0:
                newsletters.first.click()
                self.page.wait_for_load_state('networkidle')
                self.page.wait_for_timeout(2000)
                logger.info("Found and opened newsletter")
                return self.page.url

            # Alternative: try clicking on any email that contains "newsletter"
            any_newsletter = self.page.locator('text=/newsletter/i').first
            if any_newsletter.is_visible(timeout=3000):
                any_newsletter.click()
                self.page.wait_for_load_state('networkidle')
                return self.page.url

            logger.warning("Could not find newsletter")
            return None

        except Exception as e:
            logger.error(f"Failed to get newsletter: {e}")
            return None

    def get_sway_link(self) -> Optional[str]:
        """Find and return the Sway link from the newsletter."""
        logger.info("Looking for Sway link...")

        try:
            # Look for "Go to this Sway" link
            sway_link = self.page.locator('a:has-text("Go to this Sway"), a:has-text("Sway"), a[href*="sway.cloud.microsoft"]').first

            if sway_link.is_visible(timeout=5000):
                href = sway_link.get_attribute('href')
                logger.info(f"Found Sway link: {href}")
                return href

            # Try finding any Microsoft Sway URL in the page content
            content = self.page.content()
            sway_match = re.search(r'https://sway\.cloud\.microsoft\.com/[^\s"\'<>]+', content)
            if sway_match:
                logger.info(f"Found Sway URL in content: {sway_match.group()}")
                return sway_match.group()

            logger.warning("No Sway link found")
            return None

        except Exception as e:
            logger.error(f"Failed to get Sway link: {e}")
            return None

    def scrape_sway_diary_dates(self, sway_url: str) -> List[Dict]:
        """Scrape diary dates from the Sway page."""
        logger.info(f"Scraping Sway page: {sway_url}")

        events = []

        try:
            self.page.goto(sway_url)
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(3000)  # Give Sway time to fully render

            # Get all text content
            content = self.page.content()

            # Look for diary dates section
            # Sway pages often have tables or structured content

            # Try to find table rows with date information
            # Common patterns: "Friday 30th January | Red class celebration assembly | 2:45pm"

            # Method 1: Look for table cells
            rows = self.page.locator('table tr, [role="row"]').all()

            for row in rows:
                try:
                    cells = row.locator('td, [role="cell"]').all()
                    if len(cells) >= 2:
                        text = ' '.join([cell.inner_text() for cell in cells])
                        event = self._parse_event_text(text)
                        if event:
                            events.append(event)
                except:
                    continue

            # Method 2: Look for date patterns in text content
            if not events:
                text_content = self.page.locator('body').inner_text()
                events = self._extract_events_from_text(text_content)

            logger.info(f"Found {len(events)} events in Sway page")
            return events

        except Exception as e:
            logger.error(f"Failed to scrape Sway: {e}")
            return []

    def _parse_event_text(self, text: str) -> Optional[Dict]:
        """Parse a line of text into an event dictionary."""
        # Common date patterns
        date_patterns = [
            r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2})(st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)',
            r'(\d{1,2})(st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)',
        ]

        # Time patterns
        time_pattern = r'(\d{1,2}[:.]\d{2}\s*(am|pm)?|\d{1,2}\s*(am|pm)|All day|TBC)'

        date_match = None
        for pattern in date_patterns:
            date_match = re.search(pattern, text, re.IGNORECASE)
            if date_match:
                break

        if not date_match:
            return None

        time_match = re.search(time_pattern, text, re.IGNORECASE)
        time_str = time_match.group() if time_match else 'All day'

        # Extract event title (everything that's not the date or time)
        title = text
        title = re.sub(date_patterns[0], '', title, flags=re.IGNORECASE)
        title = re.sub(date_patterns[1], '', title, flags=re.IGNORECASE)
        title = re.sub(time_pattern, '', title, flags=re.IGNORECASE)
        title = re.sub(r'\s+', ' ', title).strip()
        title = re.sub(r'^[\s|,\-]+|[\s|,\-]+$', '', title)  # Clean up separators

        if not title:
            return None

        return {
            'date_text': date_match.group(),
            'title': title,
            'time': time_str,
            'raw_text': text
        }

    def _extract_events_from_text(self, text: str) -> List[Dict]:
        """Extract events from raw text content."""
        events = []
        lines = text.split('\n')

        for line in lines:
            line = line.strip()
            if len(line) > 10:  # Skip very short lines
                event = self._parse_event_text(line)
                if event:
                    events.append(event)

        return events


class EventFilter:
    """Filters and categorizes school events."""

    @staticmethod
    def is_relevant(event: Dict) -> bool:
        """Check if event is relevant for our children (YR, Y2, KS1)."""
        title = event.get('title', '').lower()
        raw = event.get('raw_text', '').lower()
        combined = f"{title} {raw}"

        # Check for include keywords
        has_include = any(kw in combined for kw in INCLUDE_KEYWORDS)

        # Check if it's ONLY for excluded year groups
        has_exclude_only = any(kw in combined for kw in EXCLUDE_ONLY_KEYWORDS)
        has_any_include = any(kw in combined for kw in INCLUDE_KEYWORDS)

        # Include if has our keywords, or exclude if only has other year groups
        if has_include:
            return True
        if has_exclude_only and not has_any_include:
            return False

        # Include school-wide events that don't specify a year group
        if not any(f'y{i}' in combined or f'year {i}' in combined for i in range(1, 7)):
            if not 'ks2' in combined:
                return True  # Likely school-wide event

        return False

    @staticmethod
    def categorize(event: Dict) -> Tuple[str, str]:
        """
        Categorize event and return (child_name, color_id).
        Returns who the event is for and the calendar color to use.
        """
        title = event.get('title', '').lower()
        raw = event.get('raw_text', '').lower()
        combined = f"{title} {raw}"

        # Check for specific year groups
        is_yr = any(kw in combined for kw in ['yr', 'reception', 'red class'])
        is_y2 = any(kw in combined for kw in ['y2', 'year 2', 'yellow class'])
        is_ks1 = 'ks1' in combined or (is_yr and is_y2)

        if is_ks1 or (is_yr and is_y2):
            return ('Both', COLOR_BOTH)
        elif is_yr:
            return ('Arvi', COLOR_ARVI)
        elif is_y2:
            return ('Rivan', COLOR_RIVAN)
        else:
            # Default to both for school-wide events
            return ('Both', COLOR_BOTH)

    @staticmethod
    def filter_events(events: List[Dict]) -> List[Dict]:
        """Filter and enhance events with categorization."""
        filtered = []

        for event in events:
            if EventFilter.is_relevant(event):
                child, color = EventFilter.categorize(event)
                event['child'] = child
                event['color_id'] = color
                filtered.append(event)

        logger.info(f"Filtered {len(events)} events to {len(filtered)} relevant events")
        return filtered


class GoogleCalendarSync:
    """Handles Google Calendar operations."""

    SCOPES = ['https://www.googleapis.com/auth/calendar']

    def __init__(self, token_json: str = None, credentials_json: str = None):
        """
        Initialize with either token JSON string or file paths.
        For GitHub Actions, pass the token as a JSON string from secrets.
        """
        self.service = None
        self._init_service(token_json, credentials_json)

    def _init_service(self, token_json: str, credentials_json: str):
        """Initialize the Google Calendar service."""
        creds = None

        # Try to get credentials from environment (for GitHub Actions)
        if token_json:
            try:
                token_data = json.loads(token_json)
                creds = Credentials.from_authorized_user_info(token_data, self.SCOPES)
            except Exception as e:
                logger.error(f"Failed to parse token JSON: {e}")

        # Try to get from file (for local development)
        if not creds and os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)

        # Refresh if expired
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                logger.info("Refreshed Google credentials")
            except Exception as e:
                logger.error(f"Failed to refresh credentials: {e}")
                raise

        if not creds or not creds.valid:
            raise ValueError("No valid Google credentials available")

        self.service = build('calendar', 'v3', credentials=creds)
        logger.info("Google Calendar service initialized")

    def get_existing_events(self, start_date: datetime, end_date: datetime) -> List[Dict]:
        """Get existing events in date range."""
        try:
            events_result = self.service.events().list(
                calendarId='primary',
                timeMin=start_date.isoformat() + 'Z',
                timeMax=end_date.isoformat() + 'Z',
                singleEvents=True,
                orderBy='startTime'
            ).execute()

            events = events_result.get('items', [])
            logger.info(f"Found {len(events)} existing events in calendar")
            return events

        except HttpError as e:
            logger.error(f"Failed to get existing events: {e}")
            return []

    def is_duplicate(self, new_event: Dict, existing_events: List[Dict]) -> bool:
        """Check if event already exists in calendar."""
        new_title = new_event.get('title', '').lower().strip()
        new_date = new_event.get('date_parsed')

        for existing in existing_events:
            existing_title = existing.get('summary', '').lower().strip()
            existing_start = existing.get('start', {})
            existing_date = existing_start.get('date') or existing_start.get('dateTime', '')[:10]

            # Check if titles are similar and dates match
            if new_date and existing_date:
                if new_date == existing_date:
                    # Fuzzy title match - check if significant words overlap
                    new_words = set(new_title.split())
                    existing_words = set(existing_title.split())
                    common = new_words & existing_words

                    if len(common) >= 2 or new_title in existing_title or existing_title in new_title:
                        logger.debug(f"Duplicate found: '{new_title}' matches '{existing_title}'")
                        return True

        return False

    def parse_date(self, date_text: str, year: int = None) -> Optional[str]:
        """Parse date text to ISO format (YYYY-MM-DD)."""
        if not year:
            year = datetime.now().year

        # Month mapping
        months = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4,
            'may': 5, 'june': 6, 'july': 7, 'august': 8,
            'september': 9, 'october': 10, 'november': 11, 'december': 12
        }

        date_text = date_text.lower()

        # Extract day and month
        day_match = re.search(r'(\d{1,2})', date_text)
        month_match = None
        for month_name, month_num in months.items():
            if month_name in date_text:
                month_match = month_num
                break

        if day_match and month_match:
            day = int(day_match.group(1))

            # If month is earlier than current month, assume next year
            if month_match < datetime.now().month:
                year += 1

            try:
                date_obj = datetime(year, month_match, day)
                return date_obj.strftime('%Y-%m-%d')
            except ValueError:
                return None

        return None

    def parse_time(self, time_text: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse time text to start and end times (HH:MM format)."""
        if not time_text or 'all day' in time_text.lower() or 'tbc' in time_text.lower():
            return None, None

        time_text = time_text.lower().strip()

        # Handle time ranges like "8:40am-9:00am" or "8:40-9:00am"
        range_match = re.search(r'(\d{1,2})[:.:]?(\d{2})?\s*(am|pm)?\s*[-–to]+\s*(\d{1,2})[:.:]?(\d{2})?\s*(am|pm)?', time_text)

        if range_match:
            start_hour = int(range_match.group(1))
            start_min = int(range_match.group(2) or 0)
            start_ampm = range_match.group(3)
            end_hour = int(range_match.group(4))
            end_min = int(range_match.group(5) or 0)
            end_ampm = range_match.group(6) or start_ampm

            # Convert to 24-hour format
            if start_ampm == 'pm' and start_hour != 12:
                start_hour += 12
            if end_ampm == 'pm' and end_hour != 12:
                end_hour += 12

            return f"{start_hour:02d}:{start_min:02d}", f"{end_hour:02d}:{end_min:02d}"

        # Handle single time like "2:45pm"
        single_match = re.search(r'(\d{1,2})[:.:]?(\d{2})?\s*(am|pm)?', time_text)

        if single_match:
            hour = int(single_match.group(1))
            minute = int(single_match.group(2) or 0)
            ampm = single_match.group(3)

            if ampm == 'pm' and hour != 12:
                hour += 12
            elif ampm == 'am' and hour == 12:
                hour = 0

            start = f"{hour:02d}:{minute:02d}"
            # Default to 1 hour duration
            end_hour = hour + 1
            end = f"{end_hour:02d}:{minute:02d}"

            return start, end

        return None, None

    def create_event(self, event: Dict) -> Optional[str]:
        """Create a calendar event and return the event ID."""
        # Parse date
        date_str = self.parse_date(event.get('date_text', ''))
        if not date_str:
            logger.warning(f"Could not parse date for event: {event.get('title')}")
            return None

        event['date_parsed'] = date_str

        # Build event title with child prefix
        child = event.get('child', 'Both')
        prefix = f"[{child}] " if child != 'Both' else "[School] "
        title = prefix + event.get('title', 'School Event')

        # Parse time
        start_time, end_time = self.parse_time(event.get('time', ''))

        # Build event body
        if start_time and end_time:
            # Timed event
            event_body = {
                'summary': title,
                'start': {
                    'dateTime': f"{date_str}T{start_time}:00",
                    'timeZone': 'Europe/London',
                },
                'end': {
                    'dateTime': f"{date_str}T{end_time}:00",
                    'timeZone': 'Europe/London',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 24 * 60},  # 1 day before
                        {'method': 'popup', 'minutes': 60},       # 1 hour before
                    ],
                },
            }
        else:
            # All-day event
            event_body = {
                'summary': title,
                'start': {
                    'date': date_str,
                },
                'end': {
                    'date': date_str,
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 24 * 60},  # 1 day before
                    ],
                },
            }

        # Add color
        if event.get('color_id'):
            event_body['colorId'] = event['color_id']

        # Add description
        event_body['description'] = f"Auto-synced from ParentMail newsletter.\n\nOriginal text: {event.get('raw_text', '')}"

        try:
            created = self.service.events().insert(
                calendarId='primary',
                body=event_body
            ).execute()

            logger.info(f"Created event: {title} on {date_str}")
            return created.get('id')

        except HttpError as e:
            logger.error(f"Failed to create event: {e}")
            return None

    def sync_events(self, events: List[Dict]) -> Tuple[int, int]:
        """
        Sync events to calendar, avoiding duplicates.
        Returns (created_count, skipped_count).
        """
        # Get existing events for the next 6 months
        start_date = datetime.now()
        end_date = start_date + timedelta(days=180)
        existing_events = self.get_existing_events(start_date, end_date)

        created = 0
        skipped = 0

        for event in events:
            # Parse date first to check for duplicates
            date_str = self.parse_date(event.get('date_text', ''))
            if date_str:
                event['date_parsed'] = date_str

            if self.is_duplicate(event, existing_events):
                logger.debug(f"Skipping duplicate: {event.get('title')}")
                skipped += 1
            else:
                event_id = self.create_event(event)
                if event_id:
                    created += 1
                    # Add to existing events to prevent duplicates within this batch
                    existing_events.append({
                        'summary': event.get('title'),
                        'start': {'date': event.get('date_parsed')}
                    })

        return created, skipped


def main():
    """Main entry point for the daily sync."""
    logger.info("=" * 50)
    logger.info("Starting ParentMail Calendar Sync")
    logger.info(f"Time: {datetime.now().isoformat()}")
    logger.info("=" * 50)

    # Check for required credentials
    if not PARENTMAIL_PASSWORD:
        logger.error("PARENTMAIL_PASSWORD environment variable not set")
        return 1

    # Get Google token from environment or file
    google_token = os.environ.get('GOOGLE_CALENDAR_TOKEN')

    events = []

    # Step 1: Scrape ParentMail
    try:
        with ParentMailScraper(PARENTMAIL_EMAIL, PARENTMAIL_PASSWORD) as scraper:
            if not scraper.login():
                logger.error("Failed to login to ParentMail")
                return 1

            newsletter_url = scraper.get_latest_newsletter()
            if not newsletter_url:
                logger.warning("Could not find latest newsletter")
                return 0  # Not an error - maybe no new newsletter

            sway_link = scraper.get_sway_link()
            if not sway_link:
                logger.warning("Could not find Sway link in newsletter")
                return 0

            events = scraper.scrape_sway_diary_dates(sway_link)

    except Exception as e:
        logger.error(f"Scraping failed: {e}")
        return 1

    if not events:
        logger.info("No events found in newsletter")
        return 0

    # Step 2: Filter events
    filtered_events = EventFilter.filter_events(events)

    if not filtered_events:
        logger.info("No relevant events after filtering")
        return 0

    # Step 3: Sync to Google Calendar
    try:
        calendar = GoogleCalendarSync(token_json=google_token)
        created, skipped = calendar.sync_events(filtered_events)

        logger.info("=" * 50)
        logger.info(f"Sync complete!")
        logger.info(f"Created: {created} events")
        logger.info(f"Skipped (duplicates): {skipped} events")
        logger.info("=" * 50)

    except Exception as e:
        logger.error(f"Calendar sync failed: {e}")
        return 1

    return 0


if __name__ == '__main__':
    exit(main())
