import http.client, urllib.request, urllib.parse, urllib.error, json, webbrowser, urllib.parse, os, hashlib, base64
from http.server import BaseHTTPRequestHandler, HTTPServer

CLIENT_ID       = "NL8IZo82T84ZCOruAZom4LlmrzkQFXPW"
CLIENT_SECRET   = "5RNNKjqAAA1szIImP0CO2IFNC6Z8OoBMQeiMKwwoxST7ntSFJhIQKVG1s1DEbLOV"
REDIRECT_URI    = "http://localhost"
PORT_NUMBER     = 80
BROWSER_CLOSE_WINDOW = '''<!DOCTYPE html> 
                        <html> 
                                <script type="text/javascript"> 
                                    function close_window() { close(); }
                                </script>
                            <body onload="close_window()"/>
                        </html>'''
server = None
urlDict2 = None
data = {}
code_verifier = base64.urlsafe_b64encode(os.urandom(64))
code_challenge = base64.urlsafe_b64encode(hashlib.sha256(code_verifier).digest()).rstrip(b'=')

#1. Logging in with access token
def read_access_token():
    try:
        with open('access_token.txt', 'r') as codeFile:
            access_token = codeFile.read()

        with open('token_type.txt', 'r') as codeFile:
            token_type = codeFile.read()
    except IOError:
        return "", ""

    return access_token, token_type

#2. If access token doesn't work, try refresh_token
def get_access_token_from_refresh_token(client_id, client_secret):
    with open('refresh_token.txt', 'r') as codeFile:
        refresh_token = codeFile.read()
    codeFile.close()

    conn = http.client.HTTPSConnection("www.wrike.com")
    urlDict = urllib.parse.urlencode({"client_id":     client_id,
                                "client_secret": client_secret,
                                "grant_type": "refresh_token",
                                "refresh_token": refresh_token, })
    headers = {"Content-type": "application/x-www-form-urlencoded", }
    conn.request("POST", "/oauth2/token", urlDict, headers)
    response = conn.getresponse()
    if response.status == http.client.UNAUTHORIZED:
        access_token, token_type = log_in()
    else:
        rjson = json.load(response)
        access_token = rjson['access_token']
    print(access_token)
    print(response)
    with open('access_token.txt', 'w') as codeFile:
        codeFile.write(access_token)
    codeFile.close()
    return access_token

#3. Logging in explicitely
def log_in():
    global server

    authorizePath = '/identity/connect/authorize'
    urlDict = urllib.parse.urlencode({"client_id"             : CLIENT_ID,
                                "response_type"         : "code",
                                "redirect_uri"          : REDIRECT_URI,
                                "scope"                 : "admin admin.brand admin.product offline_access",
                                # "scope"                 : "search_api search_api_downloadbinary",
                                "code_challenge"        : code_challenge,
                                "code_challenge_method" : "S256",
                                "state"                 : "1",
                                })

    ue = urllib.parse.urlunparse(('https',
                                'accounts.bimobject.com',
                                authorizePath,
                                '',
                                urlDict,
                                '', ))
    # print "ue " + ue
    webbrowser.open(ue)
    server = HTTPServer(('', PORT_NUMBER), myHandler)

    try:
        server.serve_forever()
    except IOError:
        pass

    urlDict2 = urllib.parse.urlencode({"client_id"        : CLIENT_ID,
                                 "client_secret"    : CLIENT_SECRET,
                                 "grant_type"       : "authorization_code",
                                 # "grant_type"       : "client_credentials_for_admin",
                                 "code"             : data['code'],
                                 "code_verifier"    : code_verifier,
                                 "redirect_uri"     : REDIRECT_URI, })

    print(urlDict2)

    headers = {"Content-type": "application/x-www-form-urlencoded", }
    conn = http.client.HTTPSConnection("accounts.bimobject.com")
    conn.request("POST", "/identity/connect/token", urlDict2, headers)
    # conn.request("GET", "/identity/connect/authorize", urlDict2, headers)
    response = conn.getresponse().read()
    print("response: " + response)

    access_token  = json.loads(response)['access_token']
    refresh_token = json.loads(response)['refresh_token']
    token_type    = json.loads(response)['token_type']

    # with open('access_token.txt', 'w') as codeFile:
    #     codeFile.write(access_token)

    with open('refresh_token.txt', 'w') as codeFile:
        codeFile.write(refresh_token)

    with open('token_type.txt', 'w') as codeFile:
        codeFile.write(token_type)

    return access_token, token_type


class myHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global data
        self.wfile.write(BROWSER_CLOSE_WINDOW)
        data = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        data = dict([(i, data[i][0]) if data[i] else (i, '') for i in data])
        print(data['code'])

        server.server_close()
        with open('code.txt', 'w') as codeFile:
            codeFile.write(data['code'])


def getResponse(token_type,
                access_token,
                inURL = "api.bimobject.com",
                inPath="/admin/v1/brands?%s",
                inHeaders = {"Content-type": "application/x-www-form-urlencoded",
                             "Authorization": ""},
                inQuery={"fields": "name, id", "pageSize": 1000},):
    inHeaders["Authorization"] = "%s %s" % (token_type, access_token,)
    conn = http.client.HTTPSConnection(inURL)
    urlDict = urllib.parse.urlencode(inQuery)
    conn.request("GET", inPath % urlDict, urlDict, inHeaders)
    response = conn.getresponse()
    print(response.status, response.reason)
    return response.read(), response.status, response.reason

#TODO try: and def
access_token, token_type = read_access_token()
print(1)
conn = http.client.HTTPSConnection("api.bimobject.com")
urlDict = urllib.parse.urlencode({"fields": "name, id",
                            "pageSize": 1000})
headers = {"Content-type": "application/x-www-form-urlencoded",
           "Authorization": token_type + " " + access_token}
# conn.request("GET", "/admin/v1/brands", urlDict, headers)
conn.request("GET", "/admin/v1/masterdata/products/filetypes", urlDict, headers)
# conn.request("GET", "/admin/v1/brands?%s" % urlDict, urlDict, headers)
# conn.request("GET", "/search/v1/products", urlDict, headers)
response = conn.getresponse()
print(response.read())
print(response.status, response.reason)
# print json.load(response)

if response.status == 401:
    access_token, token_type = log_in()
    print(2)
    conn = http.client.HTTPSConnection("api.bimobject.com")
    # urlDict = urllib.urlencode({})
    headers = {"Content-type": "application/x-www-form-urlencoded",
               "Authorization": token_type + " " + access_token}
    # conn.request("GET", "/admin/v1/brands", urlDict, headers)
    # conn.request("GET", "/admin/v1/brands?%s" % urlDict, urlDict, headers)
    conn.request("GET", "/admin/v1/masterdata/products/filetypes", urlDict, headers)
    # conn.request("GET", "/search/v1/products", urlDict, headers)
    response = conn.getresponse()
    # print json.load(response)
    print(response.read())