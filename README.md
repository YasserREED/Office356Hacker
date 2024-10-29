
# Office365Hacker
**Illicit Consent Phishing Attack Tool:** This repository hosts a tool for simulating a phishing attack where a user is tricked into granting permissions to an attacker’s Azure app. Once permissions are granted, the attacker gains access to sensitive data, such as emails, files in OneDrive, and can send emails on behalf of the user.

- This attack workflow demonstrates how attackers can leverage a legitimate Microsoft login page to gain user consent, leading to unauthorized access to user data within an organization.
- Blog Reference: <a href="https://www.alteredsecurity.com/post/introduction-to-365-stealer">Detailed Consent Grant Attack Flow</a> 


![image](https://github.com/user-attachments/assets/f5dcdad2-7564-4225-a147-b4939c2683ab)

![](https://img.shields.io/badge/Version-%20v1.0.0-blue)
![](https://img.shields.io/badge/Twitter-%20YasserREED-blue)
![](https://img.shields.io/badge/YouTube-%20YasserRED-red)


## Disclaimer

**Office365Hacker** is a security research tool intended strictly for educational, ethical research, and legal penetration testing purposes. This repository simulates phishing scenarios to raise awareness about security vulnerabilities within Office 365 and improve defensive measures. **It is not intended for, nor should it be used in, any malicious or unauthorized activity.**

Using these tools without explicit permission from the target system's owner is illegal, unethical, and can result in serious legal consequences. The authors and contributors of **Office365Hacker** assume no liability or responsibility for any misuse, damage, or harm caused by unauthorized use of this code or techniques. Please ensure all actions comply with local laws, regulations, and ethical standards.

By using **Office365Hacker**, you agree to use it responsibly, for approved security research, and in accordance with all applicable laws and ethical guidelines. 

## Phishing Attack Flow and Tool Usage

### Setup
1. **Download the Tool**

```bash
sudo git clone https://github.com/YourUsername/Office356Hacker.git
cd Office356Hacker
```
2. **Create a virtual environment to install necessary packages:**

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
deactivate
```

---

#### Phishing Attack Workflow
1. **Send Phishing Link:** Craft a link to the legitimate Microsoft login page with the attacker’s app details.
2. **Capture Authorization Code:** When a user accepts permissions, capture the authorization code from the redirect URL.
3. **Request Access Tokens:** Use the authorization code to gain access tokens for the target account.
4. **Perform Actions:** With the granted permissions, the attacker can view the inbox, download files, or send emails.

---

### `Attack Process`
#### Setup New Registry at Azure AD for Attacker
1. Access [Azure Portal Home](https://portal.azure.com/#home)

![image](https://github.com/user-attachments/assets/d20609fa-3a1e-48e2-b65c-cb8be4643599)

2. Select **Microsoft Entra ID**

![image](https://github.com/user-attachments/assets/6d50396a-221f-4ff6-aa1c-067840905c93)

3. Go to **App registrations**

![image](https://github.com/user-attachments/assets/2b6d8c50-27c8-4a9c-8947-8ce6347c30fa)

4. Click on **New registration**
 
![image](https://github.com/user-attachments/assets/066f13ac-bf30-4988-9427-73b3737a5082)

5. Add the following information using your `ngrok` link as the redirect URL
 
```console
    └─$ ngrok http 5000
```

Copy the link to use it:
- Link: `https://f214-2001-16a2-da57-9400-1d0c-cc4d-3b73-551a.ngrok-free.app`

![image](https://github.com/user-attachments/assets/5386f075-bf70-420d-a122-77dbc2f50a2b)

Add it to the Redirect URI:
- `https://f214-2001-16a2-da57-9400-1d0c-cc4d-3b73-551a.ngrok-free.app/login/authorized`

![image](https://github.com/user-attachments/assets/4b9c8be7-dafb-49ef-a0be-4bdfa7ff5731)

6. Copy Client ID, Then go to **Add certification or secret** to create a secret value
  - `Client-ID = ece26f81-4241-4e03-9207-c4eb3ec964d6`

![image](https://github.com/user-attachments/assets/e97e2a7c-1e3b-4e14-8848-bf2b5d8074b7)

6. Copy the secret value:  
  - `Secret Value`: `IR58Q~W1********************`  

![image](https://github.com/user-attachments/assets/a322283d-e7af-4026-9be3-026f8a9a94ef)

![image](https://github.com/user-attachments/assets/fa3d9576-9e6a-4074-a9c6-ddac41be1d21)

8. Add the required permissions:

**Add the following permissions:**
- User.Read
- Contacts.Read
- Mail.Read
- Mail.Send
- MailboxSettings.ReadWrite
- Files.ReadWrite.All
- User.ReadBasic.All
- Sites.Read.All
- Sites.ReadWrite.All

![image](https://github.com/user-attachments/assets/aec847e7-8777-43b6-88da-b3650e671fb6)

---

### Collected Information
```
Client ID: ece26f81-4241-4e03-9207-c4eb3ec964d6
Secret Value: IR58Q~W1********************
Redirect URL: https://f214-2001-16a2-da57-9400-1d0c-cc4d-3b73-551a.ngrok-free.app/login/authorized
```

Populate this information in `app_config.py`.

![image](https://github.com/user-attachments/assets/17a686b0-3521-4263-bdae-76175499e753)

#### Extracting Data Using the Tool
1. Run the listener to capture the token and copy the phishing URL.

```console
attacker@Machine$ python3 app.py
```

_Note: You can leave it running for additional victim tokens._

![image](https://github.com/user-attachments/assets/10362b1d-89da-442b-a0be-f5f871e04abe)
  
2. Send an email with the URL to the victim and wait for them to accept the permissions.

![image](https://github.com/user-attachments/assets/110593b1-e91c-4a62-a023-8c571311615a)

---

#### Victim Side

![image](https://github.com/user-attachments/assets/b73b44fc-b67d-4d48-9a4a-2ea72ffab408)

Once redirected to the intended site https://outlook.office.com, the attack can proceed.

![image](https://github.com/user-attachments/assets/def85d0e-f461-4303-920b-cfb8a43359f6)

3. Then You can see in CLI the attack is successful

![image](https://github.com/user-attachments/assets/3cd79f92-1b9e-470a-82c6-54a10016902c)

4. When the token is obtained, run `office365Hacker.py` to exploit the token via Microsoft Graph API

```console
attacker@Machine$ python3 office365Hacker.py --token token_data.json
```
_Note: Can handle multiple victims simultaneously._

![image](https://github.com/user-attachments/assets/021115f3-6a20-4e5d-9f8d-c3da4d0bab57)

---

#### Send Email Attack
To send an email on behalf of the victim:

```console
Office365Hacker@user1 > run send_email
```

![image](https://github.com/user-attachments/assets/38991a21-4011-45e0-94d0-fc742f039eb9)

Attacker's inbox

![image](https://github.com/user-attachments/assets/cf71a69c-cb82-4e93-8a57-91fea18211a2)

<br>

---

<p align="center"> Enjoy! :heart_on_fire: </p>
