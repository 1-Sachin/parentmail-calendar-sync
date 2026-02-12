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
import base64
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple

from playwright.sync_api import sync_playwright, Page, Browser
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import anthropic

# Email configuration
NOTIFICATION_EMAIL = os.environ.get('NOTIFICATION_EMAIL', 'sachinsharma0787@gmail.com')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')  # Gmail App Password

# Anthropic API for vision-based diary date extraction
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')

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
EXCLUDE_CLASS_ASSEMBLIES = ['orange class', 'green class', 'blue class', 'purple class', 'silver class']


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

    def handle_cookie_banner(self):
        """Handle cookie consent banner if present."""
        try:
            # Look for common cookie accept/reject buttons
            cookie_selectors = [
                'button:has-text("Accept")',
                'button:has-text("Reject")',
                'button:has-text("Accept All")',
                'button:has-text("Accept Cookies")',
                '[id*="cookie"] button',
                '[class*="cookie"] button',
                'button[id*="accept"]',
                'button[id*="reject"]',
            ]

            for selector in cookie_selectors:
                try:
                    btn = self.page.locator(selector).first
                    if btn.is_visible(timeout=2000):
                        # Prefer reject for privacy, but accept works too
                        btn.click()
                        logger.info(f"Clicked cookie banner button: {selector}")
                        self.page.wait_for_timeout(1000)
                        return True
                except:
                    continue

            return False
        except Exception as e:
            logger.warning(f"Error handling cookie banner: {e}")
            return False

    def login(self) -> bool:
        """Log into ParentMail via IRIS OAuth flow."""
        logger.info("Logging into ParentMail...")

        try:
            self.page.goto(PARENTMAIL_URL)
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(2000)

            # Handle cookie consent banner if present
            self.handle_cookie_banner()
            self.page.wait_for_timeout(1000)

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

            # Wait for potential "Stay signed in?" prompt to fully load
            logger.info("Waiting for 'Stay signed in' page to load...")
            self.page.wait_for_timeout(5000)
            self.page.wait_for_load_state('networkidle')

            self.page.screenshot(path='step6b_checking_stay_signed_in.png')

            # Log what's on the page
            page_content = self.page.locator('body').inner_text()
            logger.info(f"Page content after password submit: {page_content[:500]}")

            # Check if we're on the "Keep me signed in" page
            if 'Keep me signed in' in page_content or 'Stay signed in' in page_content:
                logger.info("Found 'Stay signed in' prompt - attempting to click...")

                clicked = False

                # Method 1: Try clicking by exact role and name
                try:
                    btn = self.page.get_by_role("button", name="Stay signed in")
                    if btn.is_visible(timeout=3000):
                        btn.click()
                        logger.info("Clicked 'Stay signed in' using get_by_role")
                        clicked = True
                        self.page.wait_for_timeout(3000)
                except Exception as e:
                    logger.info(f"get_by_role method failed: {e}")

                # Method 2: Try clicking by text content
                if not clicked:
                    try:
                        btn = self.page.get_by_text("Stay signed in", exact=True)
                        if btn.is_visible(timeout=2000):
                            btn.click()
                            logger.info("Clicked 'Stay signed in' using get_by_text")
                            clicked = True
                            self.page.wait_for_timeout(3000)
                    except Exception as e:
                        logger.info(f"get_by_text method failed: {e}")

                # Method 3: Try various CSS selectors
                if not clicked:
                    selectors = [
                        'button:has-text("Stay signed in")',
                        'button:text-is("Stay signed in")',
                        'button:has-text("Don\'t stay signed in")',
                    ]
                    for selector in selectors:
                        try:
                            btn = self.page.locator(selector).first
                            if btn.is_visible(timeout=2000):
                                btn.click()
                                logger.info(f"Clicked button with selector: {selector}")
                                clicked = True
                                self.page.wait_for_timeout(3000)
                                break
                        except Exception as e:
                            logger.debug(f"Selector {selector} failed: {e}")

                # Method 4: Find all buttons and click the right one
                if not clicked:
                    logger.info("Trying to find all buttons...")
                    buttons = self.page.locator('button').all()
                    logger.info(f"Found {len(buttons)} buttons")
                    for i, btn in enumerate(buttons):
                        try:
                            btn_text = btn.inner_text()
                            logger.info(f"Button {i}: '{btn_text}'")
                            if 'Stay signed in' in btn_text:
                                btn.click()
                                logger.info(f"Clicked button: '{btn_text}'")
                                clicked = True
                                self.page.wait_for_timeout(3000)
                                break
                        except Exception as e:
                            logger.debug(f"Button {i} error: {e}")

                if not clicked:
                    logger.error("FAILED to click 'Stay signed in' button!")
                    self.page.screenshot(path='stay_signed_in_failed.png')
            else:
                logger.info("No 'Stay signed in' prompt detected - may have been skipped")

            self.page.screenshot(path='step6b_after_stay_signed_in.png')

            # Wait for redirect back to ParentMail - may take a while
            logger.info("Waiting for redirect back to ParentMail...")

            # Wait for URL to change to ParentMail domain
            max_wait = 15  # seconds
            for i in range(max_wait):
                self.page.wait_for_timeout(1000)
                current_url = self.page.url
                logger.info(f"Redirect check {i+1}/{max_wait}: {current_url}")

                if 'pmx.parentmail.co.uk' in current_url:
                    logger.info("Redirect to ParentMail detected!")
                    break

            self.page.wait_for_load_state('networkidle')
            self.page.screenshot(path='step7_after_login_submit.png')
            logger.info(f"Step 7 - Final URL after login: {self.page.url}")

            # Check if login was successful
            if 'pmx.parentmail.co.uk' in self.page.url:
                logger.info("Redirected to ParentMail domain")

                # Handle cookie banner (may appear after redirect)
                self.handle_cookie_banner()
                self.page.wait_for_timeout(2000)

                self.page.screenshot(path='step8_parentmail_after_login.png')
                logger.info(f"ParentMail after login, URL: {self.page.url}")

                # Check if we're ACTUALLY logged in by looking at page content
                page_text = self.page.locator('body').inner_text()
                logger.info(f"Page content after login (first 500 chars): {page_text[:500]}")

                # If we landed on a valid page (emails, feed, etc.), we're logged in!
                if '/web/' in self.page.url or '/feed' in self.page.url or '/emails' in self.page.url:
                    logger.info("Login successful - landed on authenticated page!")
                    return True

                # Look for signs we're logged in (emails link, dashboard, etc.)
                if 'Emails' in page_text or 'Messages' in page_text or 'Dashboard' in page_text or 'Sachin' in page_text:
                    logger.info("Login successful - found dashboard elements!")
                    return True

                # If we see "To Register" or login page content, we're NOT logged in
                if 'To Register' in page_text or 'follow the link' in page_text.lower():
                    logger.error("Not actually logged in - still seeing registration page")
                    self.page.screenshot(path='login_not_actually_working.png')
                    return False

                # URL check - if on login page, not logged in
                if '#core/login' in self.page.url:
                    logger.error("Still on login page - login failed")
                    return False

                logger.info("Login appears successful")
                return True

            # Alternative check: look for dashboard elements
            try:
                if self.page.locator('text=Emails').is_visible(timeout=5000):
                    logger.info("Login successful - found Emails link!")
                    return True
            except:
                pass

            # Check if we're stuck on login/identity page
            if 'identity.iris.co.uk' in self.page.url or 'login' in self.page.url.lower():
                logger.error("Login failed - still on login/identity page")
                self.page.screenshot(path='login_failed.png')

                # Log page content to see if there's an error message
                page_text = self.page.locator('body').inner_text()
                logger.error(f"Login page content: {page_text[:500]}")
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

    def get_emails_list(self) -> List[Dict]:
        """Navigate to emails and get list of recent emails."""
        logger.info("Getting list of recent emails...")

        try:
            # Navigate to emails section
            self.page.goto(f"{PARENTMAIL_URL}/ui/#/messages/emails")
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(3000)

            self.page.screenshot(path='emails_list.png')

            # Get all email items in the list
            # ParentMail typically shows emails in a list/table format
            email_items = self.page.locator('[class*="message"], [class*="email"], [class*="item"], tr[class*="row"], .mail-item, .message-row').all()

            if not email_items:
                # Try alternative: look for clickable elements that might be emails
                email_items = self.page.locator('div[role="listitem"], div[role="row"], .clickable').all()

            logger.info(f"Found {len(email_items)} potential email items")

            # Log what we see on the page for debugging
            page_text = self.page.locator('body').inner_text()
            logger.info(f"Emails page content (first 1000 chars): {page_text[:1000]}")

            return email_items[:10]  # Check up to 10 most recent emails

        except Exception as e:
            logger.error(f"Failed to get emails list: {e}")
            return []

    def check_email_for_events(self, email_element) -> Optional[List[Dict]]:
        """Click on an email and check if it contains event information."""
        try:
            # Click on the email to open it
            email_element.click()
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(2000)

            # Get email content
            email_content = self.page.locator('body').inner_text()
            logger.info(f"Checking email content (first 300 chars): {email_content[:300]}")

            # Check for Sway link
            sway_link = self.get_sway_link()
            if sway_link:
                logger.info(f"Found Sway link in email: {sway_link}")
                events = self.scrape_sway_diary_dates(sway_link)
                if events:
                    return events

            # Check for date patterns directly in email (for non-Sway emails)
            events = self._extract_events_from_text(email_content)
            if events:
                logger.info(f"Found {len(events)} events directly in email")
                return events

            return None

        except Exception as e:
            logger.warning(f"Error checking email: {e}")
            return None

    def scan_all_recent_emails(self) -> List[Dict]:
        """Scan recent emails for any event information."""
        logger.info("Scanning recent emails for events...")

        all_events = []

        try:
            # Check if we're already on an emails page
            current_url = self.page.url
            if '/emails' in current_url or '/feed' in current_url or '/messages' in current_url:
                logger.info(f"Already on emails page: {current_url}")
                page_loaded = True
            else:
                page_loaded = False

            # Try different URL patterns for emails section
            email_urls = [
                f"{PARENTMAIL_URL}/web/feed-list/emails",  # This is the URL we see after login!
                f"{PARENTMAIL_URL}/#/messages/emails",
                f"{PARENTMAIL_URL}/ui/#/messages/emails",
                f"{PARENTMAIL_URL}/#messages/emails",
            ]

            if not page_loaded:
                for url in email_urls:
                    logger.info(f"Trying emails URL: {url}")
                    self.page.goto(url)
                    self.page.wait_for_load_state('networkidle')
                    self.page.wait_for_timeout(3000)

                    # Check if we got a valid page (not 404 or login page)
                    page_text = self.page.locator('body').inner_text()
                    if '404' not in page_text and 'Not Found' not in page_text and 'To Register' not in page_text:
                        logger.info(f"Successfully loaded emails page at: {url}")
                        page_loaded = True
                        break
                    else:
                        logger.warning(f"Page at {url} not valid, trying next...")

            # Handle any cookie banners that appeared
            self.handle_cookie_banner()
            self.page.wait_for_timeout(1000)

            if not page_loaded:
                # Try clicking on Emails link from current page
                logger.info("Trying to find and click Emails link...")
                try:
                    emails_link = self.page.locator('text=Emails, a:has-text("Emails"), [href*="emails"], [href*="messages"]').first
                    if emails_link.is_visible(timeout=5000):
                        emails_link.click()
                        self.page.wait_for_load_state('networkidle')
                        self.page.wait_for_timeout(2000)
                        self.handle_cookie_banner()
                        logger.info(f"Clicked Emails link, now at: {self.page.url}")
                except Exception as e:
                    logger.warning(f"Could not find Emails link: {e}")

            self.page.screenshot(path='emails_inbox.png')

            # Log page content to understand structure
            page_text = self.page.locator('body').inner_text()
            logger.info(f"Inbox page text (first 1500 chars): {page_text[:1500]}")

            # Try to find clickable email rows/items
            # Common selectors for email lists
            selectors_to_try = [
                'table tbody tr',
                '[class*="message-item"]',
                '[class*="mail-row"]',
                '[class*="email-row"]',
                'div[class*="list"] > div',
                '.message-list-item',
                '[data-testid*="message"]',
                'a[href*="message"]',
            ]

            email_elements = []
            for selector in selectors_to_try:
                elements = self.page.locator(selector).all()
                if elements and len(elements) > 0:
                    logger.info(f"Found {len(elements)} elements with selector: {selector}")
                    email_elements = elements[:5]  # Check up to 5 recent emails
                    break

            if not email_elements:
                logger.warning("Could not find email list items - trying to click first visible email link")
                # Fallback: just try to find and click any email
                first_email = self.page.locator('text=/newsletter|update|diary|dates|calendar|event/i').first
                if first_email.is_visible(timeout=3000):
                    email_elements = [first_email]

            logger.info(f"Will check {len(email_elements)} emails")

            for i, email_elem in enumerate(email_elements):
                try:
                    logger.info(f"Checking email {i+1}...")

                    # Navigate back to inbox first (if not first email)
                    if i > 0:
                        self.page.goto(f"{PARENTMAIL_URL}/ui/#/messages/emails")
                        self.page.wait_for_load_state('networkidle')
                        self.page.wait_for_timeout(2000)
                        # Re-find the element since page reloaded
                        for selector in selectors_to_try:
                            elements = self.page.locator(selector).all()
                            if elements and len(elements) > i:
                                email_elem = elements[i]
                                break

                    # Click on email
                    email_elem.click()
                    self.page.wait_for_load_state('networkidle')
                    self.page.wait_for_timeout(2000)

                    self.page.screenshot(path=f'email_{i+1}_content.png')

                    # Check for Sway link (newsletters have these - most reliable source)
                    sway_link = self.get_sway_link()
                    if sway_link:
                        logger.info(f"Found Sway link in email {i+1}")
                        events = self.scrape_sway_diary_dates(sway_link)
                        if events:
                            all_events.extend(events)
                            logger.info(f"Extracted {len(events)} events from Sway")
                    else:
                        # For non-Sway emails, use STRICT mode:
                        # Only extract events that specifically mention YR/Y2/KS1
                        # This avoids extracting junk like "Open 20th", "Closed 6th"
                        email_text = self.page.locator('body').inner_text()
                        events = self._extract_events_from_text(email_text, strict_mode=True)
                        if events:
                            all_events.extend(events)
                            logger.info(f"Extracted {len(events)} YR/Y2-specific events from email text")

                except Exception as e:
                    logger.warning(f"Error processing email {i+1}: {e}")
                    continue

            logger.info(f"Total events found across all emails: {len(all_events)}")
            return all_events

        except Exception as e:
            logger.error(f"Failed to scan emails: {e}")
            return []

    def get_latest_newsletter(self) -> Optional[str]:
        """Legacy method - now redirects to scan_all_recent_emails."""
        # This method is kept for compatibility but the main flow now uses scan_all_recent_emails
        logger.info("Finding latest newsletter...")

        try:
            # Navigate to emails section
            self.page.goto(f"{PARENTMAIL_URL}/ui/#/messages/emails")
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(2000)

            self.page.screenshot(path='newsletter_search.png')

            # Log what's on the page
            page_text = self.page.locator('body').inner_text()
            logger.info(f"Page text (first 1000 chars): {page_text[:1000]}")

            # Try multiple patterns to find emails
            patterns = [
                'text=/Parent newsletter/i',
                'text=/newsletter/i',
                'text=/diary dates/i',
                'text=/weekly update/i',
                'text=/school update/i',
            ]

            for pattern in patterns:
                try:
                    element = self.page.locator(pattern).first
                    if element.is_visible(timeout=2000):
                        element.click()
                        self.page.wait_for_load_state('networkidle')
                        self.page.wait_for_timeout(2000)
                        logger.info(f"Found and opened email matching: {pattern}")
                        return self.page.url
                except:
                    continue

            # If no specific pattern found, try clicking the first/most recent email
            try:
                first_row = self.page.locator('table tbody tr, [class*="message"], [class*="row"]').first
                if first_row.is_visible(timeout=3000):
                    first_row.click()
                    self.page.wait_for_load_state('networkidle')
                    self.page.wait_for_timeout(2000)
                    logger.info("Opened first/most recent email")
                    return self.page.url
            except:
                pass

            logger.warning("Could not find any emails to open")
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
        """
        Scrape diary dates from the Sway page.
        
        The diary dates are embedded as an IMAGE in Sway, so we can't extract text
        from HTML. Instead we:
        1. Take a full-page screenshot of the Sway page
        2. Send it to Claude's Vision API to extract the diary dates table
        3. Parse the structured response into events
        """
        logger.info(f"Scraping Sway page: {sway_url}")

        events = []

        try:
            self.page.goto(sway_url)
            self.page.wait_for_load_state('networkidle')
            self.page.wait_for_timeout(5000)  # Give Sway extra time to render

            # Scroll down to ensure diary dates image is loaded
            # Diary dates are typically towards the bottom of the newsletter
            self.page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
            self.page.wait_for_timeout(2000)
            self.page.evaluate('window.scrollTo(0, 0)')
            self.page.wait_for_timeout(1000)

            # Take a full-page screenshot
            screenshot_path = 'sway_full_page.png'
            self.page.screenshot(path=screenshot_path, full_page=True)
            logger.info(f"Saved full-page screenshot: {screenshot_path}")

            # Method 1: Try standard HTML table extraction first (in case format changes)
            rows = self.page.locator('table tr, [role="row"]').all()
            if rows:
                logger.info(f"Found {len(rows)} HTML table rows - trying text extraction")
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

            # Method 2: Use Claude Vision API to read the diary dates image
            if not events:
                logger.info("No HTML table found - using Claude Vision API to read diary dates image")
                events = self._extract_events_with_vision(screenshot_path)

            logger.info(f"Found {len(events)} events in Sway page")
            return events

        except Exception as e:
            logger.error(f"Failed to scrape Sway: {e}")
            return []

    def _extract_events_with_vision(self, screenshot_path: str) -> List[Dict]:
        """
        Use Claude Vision API to extract diary dates from a screenshot.
        The diary dates are displayed as an image/table in the Sway newsletter.
        """
        if not ANTHROPIC_API_KEY:
            logger.error("ANTHROPIC_API_KEY not set - cannot use vision extraction")
            return []

        try:
            # Read screenshot and encode as base64
            with open(screenshot_path, 'rb') as f:
                image_data = base64.b64encode(f.read()).decode('utf-8')

            logger.info("Sending screenshot to Claude Vision API...")

            client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

            message = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=4096,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": image_data
                                }
                            },
                            {
                                "type": "text",
                                "text": """Look at this school newsletter screenshot. Find the "Diary Dates" table/section.

Extract EVERY event from the diary dates table and return them as a JSON array. Each event should have:
- "date_text": the full date as shown (e.g. "Tuesday 10th February")
- "title": the event name (e.g. "KS1 Welcome Wednesday")  
- "time": the time shown (e.g. "8.40am-9.00am", "2.45pm", "All day")

For multi-day events like "Monday 16th- Friday 20th March", use the start date as date_text and include the full date range in a "date_end" field.

Return ONLY the JSON array, no other text. Example:
[
    {"date_text": "Tuesday 10th February", "title": "Safer Internet Day", "time": "All Day"},
    {"date_text": "Wednesday 11th February", "title": "KS1 Welcome Wednesday", "time": "8.40am-9.00am"}
]

If you cannot find a diary dates table, return an empty array: []"""
                            }
                        ]
                    }
                ]
            )

            # Parse the response
            response_text = message.content[0].text.strip()
            logger.info(f"Claude Vision API response (first 500 chars): {response_text[:500]}")

            # Clean up response - remove markdown code fences if present
            response_text = re.sub(r'^```json\s*', '', response_text)
            response_text = re.sub(r'\s*```$', '', response_text)
            response_text = response_text.strip()

            # Parse JSON
            raw_events = json.loads(response_text)
            logger.info(f"Parsed {len(raw_events)} events from Claude Vision API")

            # Convert to our event format
            events = []
            for raw in raw_events:
                event = {
                    'date_text': raw.get('date_text', ''),
                    'title': raw.get('title', ''),
                    'time': raw.get('time', 'All day'),
                    'raw_text': f"{raw.get('date_text', '')} {raw.get('title', '')} {raw.get('time', '')}",
                }
                if raw.get('date_end'):
                    event['date_end_text'] = raw['date_end']

                if event['title'] and event['date_text']:
                    events.append(event)
                    logger.info(f"Vision extracted: {event['title']} on {event['date_text']} at {event['time']}")

            return events

        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse Claude Vision API response as JSON: {e}")
            logger.error(f"Response was: {response_text[:1000]}")
            return []
        except Exception as e:
            logger.error(f"Claude Vision API extraction failed: {e}")
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

    def _extract_events_from_text(self, text: str, strict_mode: bool = False) -> List[Dict]:
        """
        Extract events from raw text content.

        Args:
            text: The text to extract events from
            strict_mode: If True, only extract events that mention YR/Y2/KS1 keywords
                        (used for non-Sway emails to avoid extracting junk)
        """
        events = []
        lines = text.split('\n')

        # Keywords that indicate a real school event (not just a date mention)
        event_keywords = [
            'trip', 'assembly', 'meeting', 'sports', 'day', 'week', 'concert',
            'performance', 'celebration', 'parents', 'evening', 'workshop',
            'fair', 'festival', 'party', 'treat', 'activity', 'visit', 'visitor',
            'class', 'photo', 'photographs', 'homework', 'reading', 'phonics',
            'nativity', 'harvest', 'christmas', 'easter', 'term', 'holiday',
            'inset', 'training', 'club', 'breakfast', 'after school', 'disco',
            'film', 'movie', 'bike', 'helmet', 'uniform', 'book', 'library'
        ]

        # Keywords that indicate this is NOT an event (skip these)
        skip_keywords = [
            'open', 'closed', 'hours', 'am -', 'pm -', 'available', 'contact',
            'email', 'phone', 'website', 'click here', 'sign up', 'register',
            'copyright', 'privacy', 'terms', 'unsubscribe'
        ]

        # Year group keywords for strict mode
        year_keywords = ['yr', 'y2', 'y1', 'ks1', 'reception', 'year 2', 'year 1',
                        'red class', 'yellow class']

        for line in lines:
            line = line.strip()
            if len(line) < 15:  # Skip very short lines
                continue

            line_lower = line.lower()

            # Skip lines that look like non-events
            if any(skip in line_lower for skip in skip_keywords):
                continue

            # In strict mode, require year group keywords
            if strict_mode:
                if not any(kw in line_lower for kw in year_keywords):
                    continue

            event = self._parse_event_text(line)
            if event:
                title_lower = event.get('title', '').lower()

                # Skip if title is too short or generic
                if len(event.get('title', '')) < 5:
                    continue

                # Skip if title is just a number or ordinal
                if re.match(r'^(open|closed|\d+(st|nd|rd|th)?)\s*(&|and)?\s*$', title_lower):
                    continue

                # In non-strict mode, prefer events with meaningful keywords
                # but still accept others from Sway pages
                if not strict_mode or any(kw in title_lower for kw in event_keywords):
                    events.append(event)
                elif any(kw in line_lower for kw in year_keywords):
                    # Accept if it mentions a year group even without event keywords
                    events.append(event)

        return events


class EventFilter:
    """Filters and categorizes school events."""

    @staticmethod
    def is_relevant(event: Dict) -> bool:
        """
        Check if event is relevant for our children (YR, Y2, KS1).
        
        Rules:
        - INCLUDE: Events mentioning YR, Y2, KS1, Red class, Yellow class
        - INCLUDE: Whole-school events (no year group specified) e.g. World Book Day, 
          Safer Internet Day, Parent Consultations, Parents working group
        - EXCLUDE: KS2-only events
        - EXCLUDE: Y3/Y4/Y5/Y6-only events
        - EXCLUDE: Other class assemblies (Orange, Green, Blue, Purple, Silver)
        """
        title = event.get('title', '').lower()
        raw = event.get('raw_text', '').lower()
        combined = f"{title} {raw}"

        # FIRST: Exclude other class assemblies (Orange, Green, Blue, Purple, Silver)
        if any(cls in combined for cls in EXCLUDE_CLASS_ASSEMBLIES):
            return False

        # Check for include keywords
        has_include = any(kw in combined for kw in INCLUDE_KEYWORDS)

        # Check if it's ONLY for excluded year groups
        has_exclude_only = any(kw in combined for kw in EXCLUDE_ONLY_KEYWORDS)
        has_any_include = any(kw in combined for kw in INCLUDE_KEYWORDS)

        # Include if has our keywords
        if has_include:
            return True
        # Exclude if only has other year groups
        if has_exclude_only and not has_any_include:
            return False

        # Include school-wide events that don't specify a year group
        if not any(f'y{i}' in combined or f'year {i}' in combined for i in range(1, 7)):
            if 'ks2' not in combined:
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

    def sync_events(self, events: List[Dict]) -> Tuple[int, int, List[Dict]]:
        """
        Sync events to calendar, avoiding duplicates.
        Returns (created_count, skipped_count, created_events_list).
        """
        # Get existing events for the next 6 months
        start_date = datetime.now()
        end_date = start_date + timedelta(days=180)
        existing_events = self.get_existing_events(start_date, end_date)

        created = 0
        skipped = 0
        created_events = []  # Track details of created events

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
                    created_events.append({
                        'title': event.get('title'),
                        'date': event.get('date_parsed'),
                        'time': event.get('time', 'All day'),
                        'child': event.get('child', 'Both')
                    })
                    # Add to existing events to prevent duplicates within this batch
                    existing_events.append({
                        'summary': event.get('title'),
                        'start': {'date': event.get('date_parsed')}
                    })

        return created, skipped, created_events


def send_notification_email(created_events: List[Dict]) -> bool:
    """Send email notification about newly added calendar events."""
    if not SMTP_PASSWORD:
        logger.warning("SMTP_PASSWORD not set - skipping email notification")
        return False

    if not created_events:
        logger.info("No events to notify about - skipping email")
        return True

    # Build email content
    subject = f"📅 {len(created_events)} new school event(s) added to calendar"

    # Build HTML body
    events_html = ""
    for event in created_events:
        child_emoji = "👦" if event['child'] == 'Rivan' else "👧" if event['child'] == 'Arvi' else "👨‍👩‍👧‍👦"
        events_html += f"""
        <tr>
            <td style="padding: 8px; border-bottom: 1px solid #eee;">{child_emoji} {event['child']}</td>
            <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>{event['title']}</strong></td>
            <td style="padding: 8px; border-bottom: 1px solid #eee;">{event['date']}</td>
            <td style="padding: 8px; border-bottom: 1px solid #eee;">{event['time']}</td>
        </tr>
        """

    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #2e7d32;">🏫 ParentMail Calendar Sync</h2>
        <p>The following <strong>{len(created_events)} event(s)</strong> have been added to your Google Calendar:</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
            <thead>
                <tr style="background-color: #f5f5f5;">
                    <th style="padding: 10px; text-align: left;">For</th>
                    <th style="padding: 10px; text-align: left;">Event</th>
                    <th style="padding: 10px; text-align: left;">Date</th>
                    <th style="padding: 10px; text-align: left;">Time</th>
                </tr>
            </thead>
            <tbody>
                {events_html}
            </tbody>
        </table>

        <p style="color: #666; font-size: 12px;">
            This is an automated message from your ParentMail Calendar Sync.<br>
            Events are synced daily at 6:00 AM UK time.
        </p>
    </body>
    </html>
    """

    # Plain text version
    text_body = f"ParentMail Calendar Sync\n\n{len(created_events)} new event(s) added:\n\n"
    for event in created_events:
        text_body += f"- [{event['child']}] {event['title']} - {event['date']} at {event['time']}\n"

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = NOTIFICATION_EMAIL
        msg['To'] = NOTIFICATION_EMAIL

        msg.attach(MIMEText(text_body, 'plain'))
        msg.attach(MIMEText(html_body, 'html'))

        # Connect to Gmail SMTP
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(NOTIFICATION_EMAIL, SMTP_PASSWORD)
            server.sendmail(NOTIFICATION_EMAIL, NOTIFICATION_EMAIL, msg.as_string())

        logger.info(f"Notification email sent to {NOTIFICATION_EMAIL}")
        return True

    except Exception as e:
        logger.error(f"Failed to send notification email: {e}")
        return False


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

    # Step 1: Scrape ParentMail - check ALL recent emails
    try:
        with ParentMailScraper(PARENTMAIL_EMAIL, PARENTMAIL_PASSWORD) as scraper:
            if not scraper.login():
                logger.error("Failed to login to ParentMail")
                return 1

            # Scan all recent emails for events (not just newsletters)
            events = scraper.scan_all_recent_emails()

            if not events:
                logger.info("No events found in any recent emails")
                return 0

    except Exception as e:
        logger.error(f"Scraping failed: {e}")
        return 1

    # Step 2: Filter events
    filtered_events = EventFilter.filter_events(events)

    if not filtered_events:
        logger.info("No relevant events after filtering")
        return 0

    # Step 3: Sync to Google Calendar
    created_events = []
    try:
        calendar = GoogleCalendarSync(token_json=google_token)
        created, skipped, created_events = calendar.sync_events(filtered_events)

        logger.info("=" * 50)
        logger.info(f"Sync complete!")
        logger.info(f"Created: {created} events")
        logger.info(f"Skipped (duplicates): {skipped} events")
        logger.info("=" * 50)

    except Exception as e:
        logger.error(f"Calendar sync failed: {e}")
        return 1

    # Step 4: Send email notification (only if events were created)
    if created_events:
        send_notification_email(created_events)

    return 0


if __name__ == '__main__':
    exit(main())
