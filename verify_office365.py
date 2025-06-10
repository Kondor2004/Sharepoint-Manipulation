

import sys  #Obtain information from python interpreter to deyermine which interpreter is running
import importlib.metadata #to determine package metadata

def main():
    print("Python executable", sys.executable) #show the full path of the python interpreter that's executing the code

    try:
        version = importlib.metadata.version("Office365-REST-Python-Client") #try to obtain the version of the package
    except importlib.metadata.PackageNotFoundError:
        print("ERROR: Office365-REST-Python-Client is NOT installed in this interpreter.") #if the error is simplt not found, print it
        sys.exit(1)
    except Exception as ex:
        print("ERROR: Could not determine Office365-REST-Python-Client version:", ex) # else print out the error that occured
        sys.exit(1)

    print("Office365-REST-Python-Client version:", version) #print out the version of office 365

    try:
        from office365.sharepoint.client_context import ClientContext #import teh class client context from office 365 module
        print("✅ Imported: office365.sharepoint.client_context.ClientContext")
    except ModuleNotFoundError as e:
        print("❌ Failed to import ClientContext:", e)
        sys.exit(1)

    try:
        from office365.runtime.auth.user_credential import UserCredential # try to import the class user credential 
        print("✅ Imported: office365.runtime.auth.user_credential.UserCredential")
    except ModuleNotFoundError as e:
        print("❌ Failed to import UserCredential:", e)
        sys.exit(1)
    
    print("\n All required imports suceeded")


if __name__ == "__main__":
    main()