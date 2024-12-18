import requests
from requests.auth import HTTPBasicAuth

# Konfiguracja
base_url = 'https://tufintest.centrala.bzwbk/securechangeworkflow/api/securechange/'
username = 'USERNAME'
password = 'PASSWORD'
ip_to_check = input("Wprowadź adres IP do sprawdzenia: ")

# Nagłówki
headers = {'Accept': 'application/json'}

# Pobierz tickety
response = requests.get(
    url=f'{base_url}/tickets',
    auth=HTTPBasicAuth(username, password),
    headers=headers,
    verify=False
)
tickets_response = response.json()
tickets_list = tickets_response.get('tickets', {}).get('ticket', [])

# Funkcja do wyciągania IP ze źródeł i celów
def extract_ip_addresses(data, key):
    """
    Wyciąga listę adresów IP z zagnieżdżonej struktury.
    :param data: słownik JSON z danymi
    :param key: 'sources' lub 'destinations'
    """
    return [
        entry.get('ip_address')
        for entry in data.get(key, {}).get(key[:-1], [])
        if entry.get('ip_address')
    ]

# Iteracja po ticketach
for ticket in tickets_list:
    ticket_id = ticket['id']
    
    # Pobierz szczegóły ticketu
    ticket_details = requests.get(
        url=f'{base_url}/tickets/{ticket_id}',
        auth=HTTPBasicAuth(username, password),
        headers=headers,
        verify=False
    ).json()
    
    # Iteracja po krokach (steps)
    for step in ticket_details.get('steps', {}).get('step', []):
        fields = step.get('tasks', {}).get('task', {}).get('fields', {}).get('field', [])
        
        # Szukanie zagnieżdżonego obiektu multi_access_request
        for field in fields:
            if field.get('@xsi.type') == 'multi_access_request':
                # Wyciągnięcie adresów źródłowych i docelowych
                sources = extract_ip_addresses(field, 'sources')
                destinations = extract_ip_addresses(field, 'destinations')
                
                # Sprawdzenie IP
                if ip_to_check in sources or ip_to_check in destinations:
                    print(f"Adres IP {ip_to_check} znaleziono w tickecie {ticket_id}")
