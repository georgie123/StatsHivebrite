

import requests

headers = {
    'Accept': 'application/json',
    'Authorization': 'Bearer 4022ca95b1e4a0198a31d95cc2eca87fbd25b7932379d31294198dd2d11be528',
    'Content-Type': 'application/json; charset=utf-8'
}
params = (
    ('page', '1'),
    ('per_page', '100'),
    ('order', 'id'),
    ('full_profile', 'true'),
)
response = requests.post('https://prime-journal.hivebrite.com/api/admin/v1/users', headers=headers, params=params)

print(response.text)