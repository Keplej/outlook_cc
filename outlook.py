import os
import base64
import mimetypes
from pathlib import Path
import httpx
from dotenv import load_dotenv
from pyexpat.errors import messages

from ms_graph import get_access_token, MS_GRAPH_BASE_URL

"""
Retrieve specific emails from a specific folder
"""

def search_folder(headers, folder_name='drafts'):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders'
    response = httpx.get(endpoint, headers=headers)
    response.raise_for_status()
    folders = response.json().get('value', [])

    for folder in folders:
        if folder['displayName'].lower() == folder_name.lower():
            return folder
    return None

# function for getting sub folders inside of folders
def get_sub_folders(headers, folder_id):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/childFolders'
    response = httpx.get(endpoint, headers=headers)
    response.raise_for_status()
    return response.json().get('value', [])

# simplify message retrival
def get_messages(headers, folder_id=None, fields='*', top=5, order_by='receivedDateTime', order_by_desc=True, max_results=20):
    if folder_id is None:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/messages'
    else:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/messages'

    params = {
        '$select': fields,  # fields to be returned
        '$top': min(top, max_results),
        '$orderby': f'{order_by} {'desc' if order_by_desc else 'asc'}'
    }

    messages = []
    next_link = endpoint

    # pagination url
    while next_link and len(messages) < max_results:
        response = httpx.get(next_link, headers=headers, params=params)

        if response.status_code != 200:
            raise Exception(f'Failed to retrieve emails: {response.json()}')

        json_response = response.json()
        messages.extend(json_response.get('value', []))
        next_link = json_response.get('@odata.nextLink', None)
        params = None # Clear params for subsequent requests

        if next_link and len(messages) + top > max_results:
            params = {
                '$top': max_results - len(messages)
            }

    return messages[:max_results]