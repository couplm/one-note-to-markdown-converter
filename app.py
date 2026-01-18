import argparse
import json
import os
import requests
import time
from typing import Optional
from pathlib import Path
from datetime import datetime, timedelta
from urllib.parse import urljoin
from bs4 import BeautifulSoup

GRAPH_API = "https://graph.microsoft.com/v1.0"
AUTH_ENDPOINT = "https://login.microsoftonline.com/common/oauth2/v2.0"
REQUEST_DELAY = 0.5 # Delay between requests in seconds

class OneNoteConverter:
    def __init__(self, access_token: str, year: Optional[str] = ""):
        self.access_token = access_token
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        self.year = year
    
    def _fetch_all_paginated(self, url: str):
        """Fetc all items from a paginated endpoint"""
        all_items = []
        while url:
            time.sleep(REQUEST_DELAY) # Handle rate limiting
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            data = response.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return all_items

    def get_notebooks(self):
        url = f"{GRAPH_API}/me/onenote/notebooks"
        return self._fetch_all_paginated(url)

    def get_sections(self, notebook_id: str):
        url = f"{GRAPH_API}/me/onenote/notebooks/{notebook_id}/sections"
        return self._fetch_all_paginated(url)

    def get_pages(self, section_id: str):
        url = f"{GRAPH_API}/me/onenote/sections/{section_id}/pages"
        return self._fetch_all_paginated(url)

    def _sanitize_filename(self, filename: str) -> str:
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            filename = filename.replace(char, "_")
        filename = filename[:200].strip()
        return filename if filename else "Untitled"

    def _get_date_from_page(self, page:dict) -> str:
        date = page.get("createdDateTime") or page.get("lastModifiedDateTime")
        if date:
            dt = datetime.fromisoformat(date.replace('Z', '+00:00'))
            return dt.strftime(self.year)
        return ""

    def _get_page_content(self, page_id: str) -> str:
        url = f"{GRAPH_API}/me/onenote/pages/{page_id}/content"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.text

    def _get_page_title(self, page_id: str) -> str:
        url = f"{GRAPH_API}/me/onenote/pages/{page_id}"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        data = response.json()
        return data.get("title")

    def _get_page_created(self, page_id: str) -> datetime:
        url = f"{GRAPH_API}/me/onenote/pages/{page_id}"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        data = response.json()
        return datetime.strptime(data.get("createdDateTime"), "%Y-%m-%dT%H:%M:%S.%fZ")

    def _get_page_last_modified(self, page_id: str) -> datetime:
        url = f"{GRAPH_API}/me/onenote/pages/{page_id}"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.text
    
    def html_to_markdown(self, html: str) -> str:
        soup = BeautifulSoup(html, "html.parser")
        markdown = self._parse_element(soup.body if soup.body else soup)
        return markdown.strip()

    def _parse_element(self, element) -> str:
        if element.name is None:
            return str(element)

        if element.name in ["p", "div"]:
            content = "".join(self._parse_element(child) for child in element.children)
            return content + "\n\n" if content.strip() else ""

        elif element.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            level = int(element.name[1])
            content = "".join(self._parse_element(child) for child in element.children)
            return f"{'#' * level} {content.strip()}\n\n"

        elif element.name == "strong" or element.name == "b":
            content = "".join(self._parse_element(child) for child in element.children)
            return f"**{content}**"

        elif element.name == "em" or element.name == "i":
            content = "".join(self._parse_element(child) for child in element.children)
            return f"*{content}*"

        elif element.name == "ul":
            return "".join(self._parse_element(li) for li in element.find_all("li", recursive=False))

        elif element.name == "ol":
            items = element.find_all("li", recursive=False)
            return "".join(f"{i+1}. {self._parse_element(li).strip()}\n" for i, li in enumerate(items)) + "\n"

        elif element.name == "li":
            content = "".join(self._parse_element(child) for child in element.children)
            return f"- {content.strip()}\n"

        elif element.name == "code":
            content = "".join(self._parse_element(child) for child in element.children)
            return f"`{content}`"

        elif element.name == "pre":
            content = "".join(self._parse_element(child) for child in element.children)
            return f"```\n{content}\n```\n\n"

        elif element.name == "a":
            text = "".join(self._parse_element(child) for child in element.children)
            href = element.get("href", "#")
            return f"[{text}]({href})"

        elif element.name == "img":
            src = element.get("src", "")
            alt = element.get("alt", "image")
            return f"![{alt}]({src})"

        elif element.name == "br":
            return "\n"

        elif element.name in ["meta", "style", "script"]:
            return ""

        else:
            return "".join(self._parse_element(child) for child in element.children)

    def convert_notebook(self, notebook_id: str, output_dir: str):
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # Load cache of already converted pages
        cache_file = output_path / ".conversion_cache.json"
        cache = {}
        if cache_file.exists():
            cache = json.loads(cache_file.read_text())

        sections = self.get_sections(notebook_id)
        print(f"Found {len(sections)} sections")

        for section in sections:
            section_name = section["displayName"]
            section_dir = output_path / section_name
            section_dir.mkdir(exist_ok=True)

            pages = self.get_pages(section["id"])
            
            # Sort pages by creation date if using page dates
            if self.year:
                pages.sort(key=lambda p: p.get("createdDateTime", ""))

            print(f"  Section '{section_name}': {len(pages)} pages")

            for page in pages:
                page_id = page["id"]
                page_name = page["title"]
                sanitized_name = self._sanitize_filename(page_name)
                
                # Add date prefix if requested
                date_prefix = self._get_date_from_page(page) if self.year else ""

                if date_prefix:
                    filename = f"{date_prefix}-{sanitized_name}.md"
                else:
                    filename = f"{sanitized_name}.md"

                filepath = section_dir / filename

                # Skip conversion if file was already converted
                if page_id in cache and (filepath.exists() or (section_dir / f"{sanitized_name}.md").exists()):
                    print(f"  Skipping (already converted): {page_name}")
                    continue

                print(f"  Converting: {page_name}")

                try:
                    time.sleep(REQUEST_DELAY)
                    html_content = self._get_page_content(page_id)
                    markdown = self.html_to_markdown(html_content)

                    filepath.write_text(markdown, encoding="utf-8")

                    cache[page_id] = filename
                    cache_file.write_text(json.dumps(cache, indent=2))

                except requests.exceptions.HTTPError as e:
                    if e.response.status_code == 403:
                        print(f"      Skipped (Access Denied - May be password protected)")
                    elif e.response.status_code == 404:
                        print(f"      Skipped (Page not found - May be a subpage or sync issue)")
                        print(f"      Page ID: {page_id}")
                        # Mark as cached
                        cache[page_id] = filename
                        cache_file.write_text(json.dumps(cache, indent=2))
                    else:
                        print(f"      Error: {e.response.status_code} - {e.response.text}")
                except Exception as e:
                    print(f"      Error: {e}")

        print(f"Conversion complete! Files saved to {output_path}")


def _year_():
    """Get year for file prefix"""
    print("Do you want to prefix files with the page creation date? e.g. YYYY-MM-DD (y/N)")
            
    choice = input().strip().lower()
    if choice == "y":
        print("Please select a date format")
        formats = [
            "%Y-%m-%d",
            "%m-%d-%Y",
            "%d-%m-%Y"
        ]
        for i, format in enumerate(formats):
            print(f"  {i+1}. {format}")
        choice = int(input("Select format number: ")) - 1
        return formats[choice]
    else:
        print(f"Loading...")
        return ""

def authenticate():
    """Get access token"""
    print("OneNote to Markdown Converter")
    print("=" * 50)
    print("\nTo use this tool, you need a Microsoft access token.")
    print("Visit https://developer.microsoft.com/en-us/graph/graph-explorer to get one.")
    print("1. Sign in with your Microsoft account")
    print("2. Copy the access token from the 'Access token' tab")
    print("\nPaste your access token below:")

    token = input().strip()
    return token

def main():
    parser = argparse.ArgumentParser(description="Convert OneNote notebooks to Markdown")
    parser.add_argument("--token", help="Microsoft Graph API access token")
    parser.add_argument("--notebook-id", help="Specific notebook ID to convert")
    parser.add_argument("--output", "-o", default="./onenote_output", help="Output directory")
    parser.add_argument("--list", action="store_true", help="List all notebooks")
    parser.add_argument("--year", help="Prefix files with page creation date (YYYY-MM-DD)")
    
    args = parser.parse_args()
    
    # Get token
    token = args.token or authenticate()

    year = args.year or _year_()
    
    try:
        converter = OneNoteConverter(token, year)
        
        if args.list:
            print("Available notebooks:")
            notebooks = converter.get_notebooks()
            for nb in notebooks:
                print(f"  {nb['displayName']} (ID: {nb['id']})")
        else:
            notebooks = converter.get_notebooks()
            if not notebooks:
                print("No notebooks found")
                return
            
            if args.notebook_id:
                selected = next((n for n in notebooks if n["id"] == args.notebook_id), None)
                if not selected:
                    print(f"Notebook ID {args.notebook_id} not found")
                    return
            else:
                print("Available notebooks:")
                for i, nb in enumerate(notebooks):
                    print(f"  {i+1}. {nb['displayName']}")
                choice = int(input("Select notebook number: ")) - 1
                selected = notebooks[choice]
            
            converter.convert_notebook(selected["id"], args.output)
    
    except requests.exceptions.HTTPError as e:
        print(f"API Error: {e.response.status_code} - {e.response.text}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
