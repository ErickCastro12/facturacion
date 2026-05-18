import win32cred  # pip install pywin32
 
for cred in win32cred.CredEnumerate(None, 0):
    if cred["Type"] == win32cred.CRED_TYPE_GENERIC:
        print(cred["TargetName"])

cred = win32cred.CredRead("RPA", win32cred.CRED_TYPE_GENERIC)
usuario = cred["UserName"]
password = cred["CredentialBlob"].decode("utf-16-le")
print(usuario, password)