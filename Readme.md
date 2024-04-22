# Steps to run this python script

### Install Dependencies
`open command prompt in current working directory and type`
```
pip install -r requirements.txt
```
`then press Enter`

### Create Google Console project

Follow the steps below: 

1. Create a `new project` 
2. Enable `Google Drive API`
3. `Add Credentials` to your project (Service Account)
4. Add `key` (JSON)
5. rename key to `client_key.json`
6. open client_key.json, copy `client_email` and `share` your file with that email.

You can also follow this [url](https://console.cloud.google.com/welcome?project=drive-updater-421119&supportedpurview=project)

### Run script
Run this command in command-line
```
python ./run.py
```