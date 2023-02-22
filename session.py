import requests
import fake_useragent
from const import const_url

session = requests.Session()

user = fake_useragent.UserAgent().random

header = {
    "user-agent": user
}

response = session.get(const_url, headers=header).text

cookies_dict = [
    {
        "domain": key.domain,
        "name": key.name,
        "path": key.path,
        "value": key.value
    }
    for key in session.cookies
]

session2 = requests.Session()

for cookies in cookies_dict:
    session2.cookies.set(**cookies)
