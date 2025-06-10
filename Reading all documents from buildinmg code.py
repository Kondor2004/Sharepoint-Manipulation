




import sys
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

SITE_URL = "https://utoronto.sharepoint.com/XXXX/XXXXXX" #SHAREPOINT URL
USERNAME = "XXXXXX.ca" #EMAIL
PASSWORD = "XXXXXXX" #PASSWORD

SOURCE_LIB_RELATIVE_URL = "/sites/fs-projects/Clean_Up_Test2" #LIBRARY URL LOCATION

def main():

    #ATTEMPT TO ACCESS THE SHAREPOINT__________________________________________
    
    print("1) Authenticating to Sharepoint...")
    try:
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
    except Exception as e:
        print(f"‼️ Authentication failed: {e}")
        sys.exit(1)
    #___________________________________________________________________________

    #ATTEMPT TO RETRIEVE THE LIBRARY FOLDER
    print(F"\n2) Retrieving root folder at '{SOURCE_LIB_RELATIVE_URL}'...")

    try:
        root_folder = ctx.web.get_folder_by_server_relative_url(SOURCE_LIB_RELATIVE_URL) # point at that folder
        ctx.load(root_folder) #to retrieve its properties
        ctx.execute_query() #sends the request
    except Exception as e:
        print(f"‼️ Could not load folder '{SOURCE_LIB_RELATIVE_URL}': {e}")
        sys.exit(1)
    
    #_____________________________________________________________________________

    #ATTEMPT TO Retreive the Sub-Folders under the Library________________________

    print("\n3) Enumerating direct subfolders under the library root...")

    try:
        subfolders = root_folder.folders
        ctx.load(subfolders)
        ctx.execute_query()
    except Exception as e:
        print(f"‼️ Failed to enumerate subfolders: {e}")
        sys.exit(1)
    
    if not subfolders:
        print("   → (No subfolders found under the library root.)")
        sys.exit(0)

    all_folders = []
    for folder in subfolders:
        name = folder.properties.get("Name", "<no-name>")
        url = folder.properties.get("ServerRelativeURL", "<no-url>")
        all_folders.append((name,url))

    print(f"\n4) Total top‐level folders under '{SOURCE_LIB_RELATIVE_URL}': {len(all_folders)}\n")
    for name,url in all_folders:
        print(f"   • {name}\n     URL: {url}\n")

    #____________________________________________________________________________

    PREFIX = "P005"
    matched = [(n,u) for (n,u) in all_folders if n.startswith(PREFIX)]
    print(f"\n5) Filtering for folders that start with '{PREFIX}':")
    if not matched:
        print(f"   ‼️ No folders starting with '{PREFIX}' were found among the {len(all_folders)} above.")
    else:
        print(f"   → Found {len(matched)} folder(s) starting with '{PREFIX}':\n")
        for n, u in matched:
            print(f"     • {n}   (URL: {u})")
        
    print("\n✅ Done listing all folders.")



    







if __name__ == "__main__":
    main()