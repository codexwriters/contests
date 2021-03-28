import requests
import json
from enum import Enum


attachments = []


class ApiClient:
    apiUri = 'https://api.elasticemail.com/v2'
    apiKey = 'REMOVED'

    def Request(method, url, data, attachs=None):
        data['apikey'] = ApiClient.apiKey
        if method == 'POST':
            result = requests.post(ApiClient.apiUri + url, params=data, files = attachs)
        elif method == 'PUT':
            result = requests.put(ApiClient.apiUri + url, params=data)
        elif method == 'GET':
            attach = ''
            for key in data:
                attach = attach + key + '=' + data[key] + '&'
            url = url + '?' + attach[:-1]
            result = requests.get(ApiClient.apiUri + url)

        jsonMy = result.json()

        if jsonMy['success'] is False:
            return jsonMy['error']

        return jsonMy['data']


def Send(subject, EEfrom, fromName, to, bodyHtml, bodyText, isTransactional, attachmentFiles):
    attachments = []
    for name in attachmentFiles:
        attachments.append(('attachments', open(name, 'rb')))
    print(attachments)

    return ApiClient.Request('POST', '/email/send', {
                'subject': subject,
                'from': EEfrom,
                'fromName': fromName,
                'to': to,
                'bodyHtml': bodyHtml,
                'bodyText': bodyText,
                'isTransactional': isTransactional}, attachments)


def Upload(file, name=None, expiresAfterDays=35, enforceUniqueFileName=False):
    attachments = []
    for name in file:
        attachments.append(('attachments', open(name, 'rb')))

    parameters = {
        'name': name,
        'expiresAfterDays': expiresAfterDays,
        'enforceUniqueFileName': enforceUniqueFileName}

    return ApiClient.Request('POST', '/file/upload', parameters, attachments)
