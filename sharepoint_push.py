from pathlib import Path
import msal 
import requests
import base64
import json

# Configuration

CLIENT_ID = "81b5886b-4377-43fa-8742-52005f8a3723"
CLIENT_SECRET = "AXG8Q~-vHriBJJmnW5RjXSZLgf5NzoMnBDh.UbrR"
TENANT_ID = "7c9746f6-ae85-4c56-81bb-67b2905f8ff8"

SHAREPOINT_HOSTNAME = "nsconline.sharepoint.com"
SHAREPOINT_SITE_PATH = "/sites/atc"


DOCUMENT_LIBRARY = "Wayside Deployment Packages"
DESTINATION_FOLDER = "CO"

LOCAL_FILE = Path(
    r"C:\NSC\ptc_wayside_CO_20261002135214.exe.X"
)

GRAPH_URL = "https://graph.microsoft.com/v1.0"

# Authentication

def get_access_token():
    authority = (
        f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority = authority,
        client_credential = CLIENT_SECRET
    )

    result = app.acquire_token_for_client(
        scopes = ["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise RuntimeError(
            "Token acquisition failed: "
            + str({
                k: v for k, v in result.items()
                if k != "access_token"
            })
        )
    token = result["access_token"]

    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)

    claims = json.loads(
        base64.urlsafe_b64decode(payload)
    )

    print("Token diagnostics:")
    print("Audience:", claims.get("aud"))
    print("Tenant ID matches:", claims.get("tid") == TENANT_ID)
    print("Application ID matches:", claims.get("appid") == CLIENT_ID or claims.get("azp") == CLIENT_ID)
    print("Application roles:", claims.get("roles", []))

    return token

# Graph Get Request

def graph_get(url, access_token):
    response = requests.get(
         url,
         headers = {
                 "Authorization": f"Bearer {access_token}"
             }
    )
    print(f"GET {url}")
    print(f"HTTP status: {response.status_code}")

    if not response.ok:
        raise RuntimeError(
            f"Microsoft Graph request failed:\n"
            f"{response.text}"
        )

    return response.json()

# Find SharePoint Site

def get_site(access_token):

    url = (
        f"{GRAPH_URL}/sites/"
        f"{SHAREPOINT_HOSTNAME}:"
        f"{SHAREPOINT_SITE_PATH}"
    )

    site = graph_get(
        url,
        access_token
    )

    return site

# Find Document Library 

def get_drive(site_id, access_token):
    url = (
        f"{GRAPH_URL}/sites/"
        f"{site_id}/drives"
    )

    data = graph_get(
        url,
        access_token
    )

    for drive in data["value"]:
        if drive["name"] == DOCUMENT_LIBRARY:
            return drive

# Find Destination Folder 

def get_folder(drive_id, access_token):

    url = (
        f"{GRAPH_URL}/drives/"
        f"{drive_id}/root:/"
        f"{DESTINATION_FOLDER}"
    )

    folder = graph_get(
        url,
        access_token
    )

    return folder

# Upload File 

def upload_file(
        drive_id,
        folder_id,
        access_token
):
    file_name = LOCAL_FILE.name
    file_size = LOCAL_FILE.stat().st_size

    chunk_size = 10 * 1024 * 1024

    url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/"
        f"{folder_id}:/{file_name}:/createUploadSession"
    )

    headers = {
        "Authorization":
            f"Bearer {access_token}",
        "Content-Type":
            "application/json"
    }

    payload = {
        "item": {
            "@microsoft.graph.conflictBehavior": "replace",
            "name": file_name,
        }
    }

    response = requests.post(
        url,
        headers = headers,
        json = payload,
     )

    if not response.ok:
        raise RuntimeError(
            f"Could not create upload session: "
            f"{response.status_code} {response.text}"
        )
    upload_url = response.json()["uploadUrl"]
    print("Upload session created.")

    # Attempt to send file

    offset = 0
    with open(LOCAL_FILE, "rb") as file:
        while offset < file_size:
            chunk = file.read(
                min(chunk_size, file_size - offset)
            )

            end = offset + len(chunk) - 1

            chunk_headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": (
                     f"bytes {offset}-{end}/{file_size}"
                )
            }

            response = requests.put(
                upload_url,
                headers = chunk_headers,
                data = chunk
            )

            if response.status_code in (200, 201):
                result = response.json()
                print("Upload completed successfully.")

            if response.status_code == 202:
                offset = end + 1
                continue

            raise RuntimeError(
                f"Chunk upload failed: "
                f"{response.status_code} {response.text}"
            )

    print("Upload HTTP status:", response.status_code)

    if not response.ok:
        raise RuntimeError(
            f"Upload failed:\n{response.text}"
        )

    result = response.json()
    print("Upload succeeded:", result.get("name"))
    print("Upload item ID:", result.get("id"))
    return result

    # Main

def main():
        access_token = get_access_token()

        print("\nTesting access to the ATC Sharepoint site")

        site = get_site(access_token)


        print("Site request succeeded.")
        print("Site URL:", site.get("webUrl"))
        print("Site ID:", site.get("id"))

        site = get_site(
            access_token
        )

        drive = get_drive(
            site["id"],
            access_token
        )

        folder = get_folder(
            drive["id"],
            access_token
        )

        upload_file(
            drive["id"],
            folder["id"],
            access_token
        )

if __name__ == "__main__":
            main()    