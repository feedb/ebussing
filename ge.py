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


with open("gal.txt", "r") as file1:
    for line in file1:
        print(line.strip())
        account = Account(
            primary_smtp_address=line.strip(),
            config=Configuration(
            service_endpoint="https://owa.server.com/EWS/Exchange.asmx",
            credentials=Credentials('domain\\user', 'password'),
        ),
        autodiscover=False,
        access_type=DELEGATE
        )
        try:
            print(account.root.tree())
        except:
            pass
