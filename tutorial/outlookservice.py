import requests
import uuid
import json

graph_endpoint = 'https://graph.microsoft.com/v1.0{0}'


def get_me(access_token):
    get_me_url = graph_endpoint.format('/me')

    # Use OData query parameters to control the results
    #  - Only return the displayName and mail fields
    query_parameters = {'$select': 'displayName,mail'}

    r = make_api_call('GET', get_me_url, access_token, "", parameters = query_parameters)

    if r.status_code == requests.codes.ok:
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)


def get_my_messages(access_token, user_email):
    get_messages_url = graph_endpoint.format('/me/mailfolders/inbox/messages')

    query_parameters = {'$top': '10',
                        '$select': 'receivedDateTime,subject,from',
                        '$orderby': 'receivedDateTime DESC'}

    r = make_api_call('GET', get_messages_url, access_token, user_email, parameters = query_parameters)

    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)


def get_shift_reports(messages, access_token, user_email):
    message_ids = [x['id'] for x in messages['value'] if "shift report" in str(x['subject']).lower()]

    attachment_ids = get_attachment_ids(access_token, message_ids, user_email)

    attachments = []
    for attachment_id in attachment_ids:
        get_messages_url = graph_endpoint.format('/me/messages/{}/attachments/{}'.format(attachment_id[0], attachment_id[1]))

        # Use OData query parameters to control the results
        #  - Only first 10 results returned
        #  - Only return the ReceivedDateTime, Subject, and From fields
        #  - Sort the results by the ReceivedDateTime field in descending order
        query_parameters = {'$top': '30'}

        r = make_api_call('GET', get_messages_url, access_token, user_email, parameters=query_parameters)
        if (r.status_code == requests.codes.ok):
            attachments.append(r.json())
            with open("test.xls","w") as outputfile:
                outputfile.write(r.json()['contentBytes'])
            return r.json()
        else:
            return "{0}: {1}".format(r.status_code, r.text)
    return attachments


def get_attachment_ids(access_token, message_ids, user_email):
    attachment_ids = []
    for message_id in message_ids:
        get_messages_url = graph_endpoint.format('/me/messages/{}/attachments'.format(message_id))

        # Use OData query parameters to control the results
        #  - Only first 10 results returned
        #  - Only return the ReceivedDateTime, Subject, and From fields
        #  - Sort the results by the ReceivedDateTime field in descending order
        query_parameters = {'$top': '30'}

        r = make_api_call('GET', get_messages_url, access_token, user_email, parameters=query_parameters)
        if (r.status_code == requests.codes.ok):
            attachment_ids.append([message_id, r.json()['value'][-1]['id']])
        else:
            return "{0}: {1}".format(r.status_code, r.text)
    return attachment_ids


def get_my_oakland_messages(access_token, user_email):
    personal_folders_id = 'AQMkADAwATZiZmYAZC05ZjJjLTIyAGNhLTAwAi0wMAoALgAAAxbzfaEKXzNKuLDwt60CJksBAIu3Tqn4wSpEg5SgDcjE_BAAAAAUstlVAAAA'
    oaklandglass_folder_id = 'AQMkADAwATZiZmYAZC05ZjJjLTIyAGNhLTAwAi0wMAoALgAAAxbzfaEKXzNKuLDwt60CJksBAIu3Tqn4wSpEg5SgDcjE_BAAAAIBWwAAAA=='
    get_messages_url = graph_endpoint.format('/me/mailfolders/%s/childFolders/%s/messages' % (personal_folders_id, oaklandglass_folder_id))

    # Use OData query parameters to control the results
    #  - Only first 10 results returned
    #  - Only return the ReceivedDateTime, Subject, and From fields
    #  - Sort the results by the ReceivedDateTime field in descending order
    query_parameters = {'$top': '30'}

    r = make_api_call('GET', get_messages_url, access_token, user_email, parameters = query_parameters)
    if (r.status_code == requests.codes.ok):
        messages = r.json()
        shift_report_messages = get_shift_reports(messages, access_token, user_email)
        return shift_report_messages
    else:
        return "{0}: {1}".format(r.status_code, r.text)


# Generic API Sending
def make_api_call(method, url, token, user_email, payload = None, parameters = None):
    # Send these headers with all API calls
    headers = { 'User-Agent' : 'python_tutorial/1.0',
                'Authorization' : 'Bearer {0}'.format(token),
                'Accept' : 'application/json',
                'X-AnchorMailbox' : user_email }

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = { 'client-request-id' : request_id,
                        'return-client-request-id' : 'true' }

    headers.update(instrumentation)

    response = None

    if (method.upper() == 'GET'):
        response = requests.get(url, headers = headers, params = parameters)
    elif (method.upper() == 'DELETE'):
        response = requests.delete(url, headers = headers, params = parameters)
    elif (method.upper() == 'PATCH'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.patch(url, headers = headers, data = json.dumps(payload), params = parameters)
    elif (method.upper() == 'POST'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.post(url, headers = headers, data = json.dumps(payload), params = parameters)

    return response