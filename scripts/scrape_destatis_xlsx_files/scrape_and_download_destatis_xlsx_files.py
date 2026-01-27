"""
Destatis Excel Scraper

This script automates the process of searching for, scraping, and downloading all Excel (.xlsx) files
from the German Federal Statistical Office (Destatis) website. It uses configurable query parameters,
crawls through paginated results, and logs downloaded file URLs and names in a CSV file.
"""

import os
import time
import requests
import csv
import argparse
from dataclasses import dataclass
from typing import List, Set
from urllib.parse import urljoin, urlsplit, urlparse, urlunparse, parse_qsl, urlencode
from lxml import etree
from html import unescape

BASE_URL: str = "https://www.destatis.de"
FORM_URL: str = "SiteGlobals/Forms/Suche/Servicesuche_Formular.html"
DOWNLOAD_FOLDER: str = os.path.join(
    "scripts", "scrape-destatis-xlsx-files", "downloaded"
)
CSV_FILE: str = os.path.join(
    "scripts", "scrape-destatis-xlsx-files", "downloaded_links.csv"
)


@dataclass
class FileRecord:
    """
    Data class to store information about a downloaded file.

    Attributes:
        url (str): The URL from which the file was downloaded.
        file_name (str): The name of the downloaded file.
    """

    url: str
    file_name: str

    def to_list(self) -> List[str]:
        """
        Convert the FileRecord instance into a list of strings.

        Returns:
            List[str]: A list containing the URL and file name.
        """
        return [self.url, self.file_name]

    @classmethod
    def from_list(cls, row: List[str]) -> "FileRecord":
        """
        Create a FileRecord instance from a list of strings.

        Args:
            row (List[str]): A list where the first element is the URL and the second is the file name.

        Returns:
            FileRecord: A new instance of FileRecord.
        """
        return cls(url=row[0], file_name=row[1])


def save_file_record_to_csv(record: FileRecord, csv_file: str) -> None:
    """
    Append a FileRecord to a CSV file. Creates the file and writes headers if it doesn't exist.

    Args:
        record (FileRecord): The file record to save.
        csv_file (str): The path to the CSV file.
    """
    file_exists: bool = os.path.exists(csv_file)
    with open(csv_file, mode="a" if file_exists else "w", newline="") as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(["URL", "File Name"])
        writer.writerow(record.to_list())


def load_file_records_from_csv(csv_file: str) -> List[FileRecord]:
    """
    Load FileRecord objects from a CSV file.

    Args:
        csv_file (str): The path to the CSV file.

    Returns:
        List[FileRecord]: A list of FileRecord instances read from the CSV file.
    """
    file_records: List[FileRecord] = []
    if not os.path.exists(csv_file):
        return file_records

    with open(csv_file, mode="r", newline="") as file:
        reader = csv.reader(file)
        next(reader, None)  # Skip header
        for row in reader:
            file_records.append(FileRecord.from_list(row))
    return file_records


def create_folder_if_not_exists(folder: str) -> None:
    """
    Create a folder if it does not exist.

    Args:
        folder (str): The path to the folder.
    """
    if not os.path.exists(folder):
        os.makedirs(folder)


def download_file(url: str, folder: str, file_name: str) -> None:
    """
    Download a file from a given URL and save it to a specified folder.

    Args:
        url (str): The URL of the file.
        folder (str): The folder where the file should be saved.
        file_name (str): The name of the file.
    """
    parsed_url = urlsplit(url)
    clean_file_name: str = os.path.basename(parsed_url.path)
    file_path: str = os.path.join(folder, clean_file_name)

    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.HTTPError as http_err:
        if response.status_code == 404:
            print(f"File not found (404): {url}")
        else:
            print(f"HTTP error occurred: {http_err}")
        return
    except Exception as err:
        print(f"Other error occurred: {err}")
        return

    with open(file_path, "wb") as file:
        file.write(response.content)


def extract_xlsx_links(html: str) -> Set[str]:
    """
    Extract all Excel (.xlsx) file links from the HTML content.

    Args:
        html (str): The HTML content as a string.

    Returns:
        Set[str]: A set of absolute URLs for .xlsx files.
    """
    tree = etree.HTML(html)
    relative_links: List[str] = tree.xpath("//a[@href[contains(.,'.xlsx')]]/@href")
    absolute_links: Set[str] = {
        unescape(urljoin(BASE_URL, link)) for link in relative_links
    }
    return absolute_links


def get_next_page_url(html: str) -> str:
    """
    Retrieve the URL for the next page from the HTML content.

    Args:
        html (str): The HTML content as a string.

    Returns:
        str: The URL for the next page if found, else an empty string.
    """
    xpath: str = "//a[@class='forward button']/@href"
    tree = etree.HTML(html)
    next_page: List[str] = tree.xpath(xpath)
    return next_page[0] if next_page else ""


def get_file_name_from_url(url: str) -> str:
    """
    Extract the file name from a URL.

    Args:
        url (str): The URL from which to extract the file name.

    Returns:
        str: The file name extracted from the URL.
    """
    parsed_url = urlparse(url)
    file_name: str = os.path.basename(parsed_url.path)
    return file_name


def check_for_duplicate_filenames(file_records: List[FileRecord]) -> bool:
    """
    Check for duplicate file names among a list of FileRecord objects and print warnings if found.

    Args:
        file_records (List[FileRecord]): The list of FileRecord objects.

    Returns:
        bool: True if duplicates are found, False otherwise.
    """
    filenames: Set[str] = set()
    duplicates: Set[str] = set()

    for record in file_records:
        if record.file_name in filenames:
            duplicates.add(record.file_name)
        else:
            filenames.add(record.file_name)

    if duplicates:
        print(f"Warning: Duplicate filenames found: {duplicates}")
        return True
    return False


def canonicalize_url(url: str) -> str:
    """
    Canonicalize the URL by removing all query parameters except for '__blob=publicationFile'.

    Args:
        url (str): The URL to canonicalize.

    Returns:
        str: The canonicalized URL with only the __blob parameter (if present and equal to 'publicationFile').
    """
    parsed = urlparse(url)
    query_params = parse_qsl(parsed.query)
    filtered_params = [
        (k, v) for k, v in query_params if k == "__blob" and v == "publicationFile"
    ]
    new_query = urlencode(filtered_params)
    return urlunparse(parsed._replace(query=new_query))


def scrape_urls(
    queries: List[str], cached_urls: List[FileRecord], crawl_delay: float
) -> List[FileRecord]:
    """
    Scrape XLSX file URLs for each query string and return new FileRecord objects.
    URLs are canonicalized to only keep the __blob parameter if it equals 'publicationFile'.

    Args:
        queries (List[str]): A list of query strings (years or text queries).
        cached_urls (List[FileRecord]): Previously scraped FileRecords to avoid duplicates.
        crawl_delay (float): The delay in seconds between requests.

    Returns:
        List[FileRecord]: A list of new FileRecord objects with unique, canonicalized URLs.
    """
    new_file_records: List[FileRecord] = []
    cached_url_set: Set[str] = {record.url for record in cached_urls}
    absolute_form_url: str = urljoin(BASE_URL, FORM_URL)

    for query in queries:
        url: str = (
            f"{absolute_form_url}?documentType_=publication&resultsPerPage=100&templateQueryString={query}"
        )
        while True:
            print(f"Fetching page: {url}")
            html: str = requests.get(url).text
            links: Set[str] = extract_xlsx_links(html)
            new_links: List[str] = [
                link for link in links if link not in cached_url_set
            ]

            for link in new_links:
                raw_url = unescape(urljoin(BASE_URL, link))
                full_url: str = canonicalize_url(raw_url)
                file_name: str = get_file_name_from_url(full_url)
                new_record = FileRecord(full_url, file_name)
                new_file_records.append(new_record)
                save_file_record_to_csv(new_record, CSV_FILE)
                cached_url_set.add(full_url)

            next_url: str = get_next_page_url(html)
            if not next_url:
                break
            url = urljoin(BASE_URL, next_url)
            time.sleep(crawl_delay)

    return new_file_records


def download_files(file_records: List[FileRecord], crawl_delay: float) -> None:
    """
    Download files based on the list of FileRecord objects,.

    Args:
        file_records (List[FileRecord]): A list of FileRecord objects to download.
        crawl_delay (float): The delay in seconds between download requests.
    """
    for record in file_records:
        unique_file_name: str = get_file_name_from_url(record.file_name)
        file_path: str = os.path.join(DOWNLOAD_FOLDER, unique_file_name)

        if os.path.exists(file_path):
            print(f"File {unique_file_name} already exists.")
            continue
        else:
            print(f"Downloading file: {unique_file_name}")

        download_file(record.url, DOWNLOAD_FOLDER, unique_file_name)


def main() -> None:
    """
    Main function to parse command-line arguments, scrape URLs based on provided queries,
    warn on duplicate filenames, and download files.
    A new flag '--skip_scraping' allows skipping the scraping process and directly downloading
    files from the CSV log.
    """
    parser = argparse.ArgumentParser(
        description="Scrape and download XLSX files from Destatis."
    )
    parser.add_argument(
        "--crawl_delay",
        type=float,
        default=30,
        help="Crawl delay in seconds (default is 30, as specified in Destatis' robots.txt)",
    )
    parser.add_argument(
        "--queries",
        nargs="*",
        default=None,
        help="List of query strings. If not provided, defaults to years 2000-2025 and ['Bericht', 'Zensus'].",
    )
    parser.add_argument(
        "--skip_scraping",
        action="store_true",
        help="If set, skip scraping and download files from the CSV log only.",
    )
    args = parser.parse_args()

    crawl_delay: float = args.crawl_delay

    create_folder_if_not_exists(DOWNLOAD_FOLDER)
    cached_file_records: List[FileRecord] = load_file_records_from_csv(CSV_FILE)

    if args.skip_scraping:
        print("Skipping scraping. Downloading files listed in CSV log.")
        download_files(cached_file_records, crawl_delay)
        return

    if args.queries is not None and len(args.queries) > 0:
        queries: List[str] = args.queries
    else:
        queries = list(map(str, range(2000, 2026))) + ["Bericht", "Zensus"]

    new_file_records: List[FileRecord] = scrape_urls(
        queries, cached_file_records, crawl_delay
    )
    all_file_records: List[FileRecord] = cached_file_records + new_file_records

    # Warn if duplicates are found, but continue processing
    check_for_duplicate_filenames(all_file_records)

    download_files(all_file_records, crawl_delay)


if __name__ == "__main__":
    main()
