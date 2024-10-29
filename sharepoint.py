import requests
import json
import math

tenant_id = '<TENANT_ID>'
client_id = '<CLIENT_ID>'
client_secret = '<CLIENT_SECRET>'


# The Sharepoint and Onedrive base URL
# The sharepoint is usually https://xxxx.sharepoint.com
# The onedrive is usually   https://xxxx-my.sharepoint.com
info = {
    'onedrive': '<ONEDRIVE>',
    'sharepoint': '<SHAREPOINT>'
}

def get_all_sites(host, access_token):
    row_limit = 1000
    url = f"{host}/_api/search/query?querytext='contentclass:STS_Site'&RowLimit={row_limit}&startrow="
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }
    have_res = True
    i = 0
    while(have_res):
        response = requests.get(url+'{}'.format(i*row_limit), headers=headers)
        response_json = response.json()
        have_res = response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['RowCount']
        i += 1

        for result in response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']:
            title = None
            path = None
            desc = None
            for cell in result['Cells']['results']:
                if cell['Key'] == 'Title':
                    title = cell['Value']
                elif cell['Key'] == 'Path':
                    path = cell['Value']
                elif cell['Key'] == 'Description':
                    desc = cell['Value']
            try:
                print(f'{title} | {desc} | {path}')
            except UnicodeEncodeError:
                print(f'{path}')


def get_access_tokens(tenant_id, client_id, client_secret, host):
    onedrive_host = "{}-my.{}".format(
        host.split('.')[0],
        '.'.join(host.split('.')[1:])
    )

    url = f'https://accounts.accesscontrol.windows.net/{tenant_id}/tokens/OAuth/2'
    onedrive_scope = f'00000003-0000-0ff1-ce00-000000000000/{onedrive_host}@{tenant_id}'
    sharepoint_scope = f'00000003-0000-0ff1-ce00-000000000000/{host}@{tenant_id}'
    
    data = {
        'grant_type':'client_credentials',
        'client_id': f'{client_id}@{tenant_id}',
        'client_secret': client_secret,
        'resource': sharepoint_scope
    }

    req = requests.post(url, data=data)
    try:
        sharepoint_token = req.json()['access_token']
    except KeyError:
        raise ValueError("Failed to retrieve the sharepoint token : \n{}".format(req.text))

    data['resource'] = onedrive_scope
    req = requests.post(url, data=data)
    try:
        onedrive_token = req.json()['access_token']
    except KeyError:
        raise ValueError("Failed to retrieve the onedrive token : \n{}".format(req.text))

    return {
        'onedrive': onedrive_token,
        'sharepoint': sharepoint_token
    }

def search_site(info, folder, searchtext):
    if folder.startswith('/personal/'):
        host = info['onedrive']
        access_token = info['access_token']['onedrive']
    else:
        host = info['sharepoint']
        access_token = info['access_token']['sharepoint']

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }
    folder_split = folder[1:].split('/')
    site_type, site_name = folder_split[:2]

    have_res = True
    i = 0
    row_limit = 1000
    url = f"{host}/{site_type}/{site_name}/_api/search/query?querytext='{searchtext} AND Path:{host}/{site_type}/{site_name}'&RowLimit={row_limit}&startrow="
    file_list = []
    while(have_res):
        response = requests.get(url+'{}'.format(i*row_limit), headers=headers)
        try:
            response_json = response.json()
        except:
            have_res = False
            continue

        have_res = response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['RowCount']
        i += 1

        for result in response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']:
            description = None
            path = None
            date = None
            size = None
            for cell in result['Cells']['results']:
                if cell['Key'] == 'HitHighlightedSummary':
                    description = cell['Value']
                elif cell['Key'] == 'Path':
                    path = '/' + '/'.join(cell['Value'].split('/')[3:])
                elif cell['Key'] == 'LastModifiedTime':
                    try:
                        date = cell['Value'].split('T')[0]
                    except:
                        date = 'NaN'
                elif cell['Key'] == 'Size':
                    size = cell['Value']

            file_list.append({
                'description': description,
                'path': path,
                'date': date,
                'size': size
            })
    return file_list


def search_site_all(info, searchtext, filter=''):
    host = info['sharepoint']
    access_token = info['access_token']['sharepoint']

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }

    have_res = True
    i = 0
    row_limit = 1000
    if filter != '':
        url_filter = f"refinementfilters='{filter}'&"
    else:
        url_filter = ''
    if searchtext != '':
        url_search = f"querytext='{searchtext}'&"
    else:
        url_search = ''

    url = f"{host}/_api/search/query?{url_search}{url_filter}RowLimit={row_limit}&startrow="
    print(url)  
    file_list = []
    while(have_res):
        response = requests.get(url+'{}'.format(i*row_limit), headers=headers)
        try:
            response_json = response.json()
        except:
            have_res = False
            continue

        have_res = response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['RowCount']
        i += 1

        for result in response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']:
            description = None
            path = None
            date = None
            size = None
            for cell in result['Cells']['results']:
                if cell['Key'] == 'HitHighlightedSummary':
                    description = cell['Value']
                elif cell['Key'] == 'Path':
                    path = '/' + '/'.join(cell['Value'].split('/')[3:])
                elif cell['Key'] == 'LastModifiedTime':
                    try:
                        date = cell['Value'].split('T')[0]
                    except:
                        date = 'NaN'
                elif cell['Key'] == 'Size':
                    size = cell['Value']

            file_list.append({
                'description': description,
                'path': path,
                'date': date,
                'size': size
            })

    host = info['onedrive']
    access_token = info['access_token']['onedrive'] 
    have_res = True    
    while(have_res):
        response = requests.get(url+'{}'.format(i*row_limit), headers=headers)
        try:
            response_json = response.json()
        except:
            have_res = False
            continue

        have_res = response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['RowCount']
        i += 1

        for result in response_json['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']:
            description = None
            path = None
            date = None
            size = None
            for cell in result['Cells']['results']:
                if cell['Key'] == 'HitHighlightedSummary':
                    description = cell['Value']
                elif cell['Key'] == 'Path':
                    path = '/' + '/'.join(cell['Value'].split('/')[3:])
                elif cell['Key'] == 'LastModifiedTime':
                    try:
                        date = cell['Value'].split('T')[0]
                    except:
                        date = 'NaN'
                elif cell['Key'] == 'Size':
                    size = cell['Value']

            file_list.append({
                'description': description,
                'path': path,
                'date': date,
                'size': size
            })
    return file_list


def get_folder(info, folder):
    if folder.startswith('/personal/'):
        host = info['onedrive']
        access_token = info['access_token']['onedrive']
    else:
        host = info['sharepoint']
        access_token = info['access_token']['sharepoint']

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }
    folder_split = folder[1:].split('/')
    site_type, site_name = folder_split[:2]
    encoded_folder = '%20'.join(folder.split(' '))
    encoded_folder = '%27%27'.join(encoded_folder.split("'"))
    url = f'{host}/{site_type}/{site_name}/_api/Web/GetFolderByServerRelativeUrl(\'{encoded_folder}\')/Folders'
    req = requests.get(url, headers=headers)
    try:
        response = req.json()
    except:
        raise ValueError("Failed to get the sharepoint response for {}: \n".format(url, req.text))
    
    folders = []
    for elt in response['d']['results']:
        folders.append({
            'name': elt['Name'],
            'link': elt['ServerRelativeUrl'],
            'time': elt['TimeLastModified'].split('T')[0]
        })
    return folders


def get_files(info, folder):
    if folder.startswith('/personal/'):
        host = info['onedrive']
        access_token = info['access_token']['onedrive']
    else:
        host = info['sharepoint']
        access_token = info['access_token']['sharepoint']

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }

    folder_split = folder[1:].split('/')
    site_type, site_name = folder_split[:2]
    encoded_folder = '%20'.join(folder.split(' '))
    encoded_folder = '%27%27'.join(encoded_folder.split("'"))
    url = f'{host}/{site_type}/{site_name}/_api/Web/GetFolderByServerRelativeUrl(\'{encoded_folder}\')/Files'
    req = requests.get(url, headers=headers)
    try:
        response = req.json()
    except:
        raise ValueError("Failed to get the sharepoint response for {}: \n".format(url, req.text))
    
    folders = []
    for elt in response['d']['results']:
        folders.append({
            'name': elt['Name'],
            'link': elt['ServerRelativeUrl'],
            'time': elt['TimeLastModified'].split('T')[0],
            'size': elt['Length']
        })
    return folders

def download_file(info, folder):
    if folder.startswith('/personal/'):
        host = info['onedrive']
        access_token = info['access_token']['onedrive']
    else:
        host = info['sharepoint']
        access_token = info['access_token']['sharepoint']

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose'
    }
    folder_split = folder[1:].split('/')
    site_type, site_name = folder_split[:2]
    filename = folder_split[-1]
    encoded_folder = '%20'.join(folder.split(' '))
    encoded_folder = '%27%27'.join(encoded_folder.split("'"))
    url = f'{host}/{site_type}/{site_name}/_api/Web/GetFileByServerRelativePath(decodedurl=\'{encoded_folder}\')/$Value'
    req = requests.get(url, headers=headers)
    if req.status_code == 200:
        with open(filename, "wb") as file:
            file.write(req.content)
    else:
        raise ValueError(f"Failed to download the file {url}: {req.text}")
    


info['access_token'] = get_access_tokens(tenant_id, client_id, client_secret, info['sharepoint'].split('://')[-1])
current_directory = ''
while(1):
    try:
        if current_directory.endswith('/'):
            current_directory = current_directory[:-1]
        raw_cmd = input(f"{current_directory} >> ")
        print('\n')
        try:
            cmd = raw_cmd.split(' ')[0]
            content = ' '.join(raw_cmd.split(' ')[1:])
        except:
            print(f"Unknown command {cmd}")
            continue
        if cmd == 'cd':
            if content == '..':
                current_directory = '/'.join(current_directory.split('/')[:-1])
            elif content.startswith('/'):
                current_directory = content
            else:
                current_directory += f'/{content}'
        elif cmd == 'ls':
            if 'content' == '.':
                link = current_directory
            else:
                if content.startswith('/'):
                    link = content
                else:
                    link = f"{current_directory}/{content}"

            folders = get_folder(info, link)
            files = get_files(info, link)
            for elt in folders:
                print(f"{elt['time']} - (dir) {elt['name']}")

            for elt in files:
                print(f"{elt['time']} - {elt['name']} ({math.floor(int(elt['size'])/(1024*1024))}MB)")
        elif cmd == 'get':
            if content.startswith('/'):
                download_file(info, content)
            else:
                download_file(info, f"{current_directory}/{content}")
        elif cmd == 'search':
            result = search_site(info, '/'.join(current_directory.split('/')[:3]), content)
            for elt in result:
                print(f"{elt['date']} - {elt['path']} ({math.floor(int(elt['size'])/(1024*1024))}MB)\n Preview : {elt['description']}")
                print('\n')
        elif cmd == 'search_all':
            result = search_site_all(info, content)
            for elt in result:
                print(f"{elt['date']} - {elt['path']} ({math.floor(int(elt['size'])/(1024*1024))}MB)\n Preview : {elt['description']}")
                print('\n')
        elif cmd == 'exit':
            break
        else:
            print(f"The command {cmd} is not defined")
    except Exception as e:
        print(f"Unkown error...: {e}")
    print('\n')
