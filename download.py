from exchangelib import Credentials, Account, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
from exchangelib import DELEGATE, IMPERSONATION, Account
from exchangelib.services import FindPeople, GetPersona
from exchangelib.items import Contact, DistributionList, Persona
from exchangelib.fields import FieldPath
from exchangelib import Q
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


with open("emails.txt", "r") as file1:
    for line in file1:
        name = line.strip()
        success = False 
        while not success:
            try:
                print(line.strip())
                account = Account(
                    primary_smtp_address=line.strip(),
                    config=Configuration(
                    service_endpoint="https://server/EWS/Exchange.asmx",
                    credentials=Credentials('domain\\user', 'pass'),
                    max_connections=10,
                ),
                autodiscover=False,
                access_type=DELEGATE
                )
                import uuid
                qs = account.inbox.all()
                qs.page_size = 1000  # Number of IDs for FindItem to get per page
                qs.chunk_size = 10  # Number of full items for GetItem to request per call
                for msg in qs:
                    with open('out\%s.%s.inbox.eml' % (uuid.uuid4(),name), "wb") as f:
                        f.write(msg.mime_content)
                success = True
            except:
                pass
