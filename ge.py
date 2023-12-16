from exchangelib import Credentials, Account, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
from exchangelib import DELEGATE, IMPERSONATION, Account
from exchangelib.services import FindPeople, GetPersona
from exchangelib.items import Contact, DistributionList, Persona
from exchangelib.fields import FieldPath
from exchangelib import Q
import sys
import argparse
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

parser = argparse.ArgumentParser(description='MS Exchange email folder access abuse.')
parser.add_argument('-i', '--inputfile',help='File with emails')
parser.add_argument('-d', '--domain',  type=str,help='Windows AD name.')
parser.add_argument('-s', '--server',  type=str, help='URL to Exchange Server.')
parser.add_argument('-u', '--username',  type=str, help='Username.')
parser.add_argument('-p', '--password', type=str, help='Password.')

args = parser.parse_args()

def main():
    if args.inputfile is not None:
        with open(args.inputfile, "r") as file1:
            for line in file1:
                print(line.strip())
                account = Account(
                    primary_smtp_address=line.strip(),
                    config=Configuration(
                    service_endpoint='https://'+args.server+'/EWS/Exchange.asmx',
                    credentials=Credentials(args.domain+'\\'+args.username, args.password),
                ),
                autodiscover=False,
                access_type=DELEGATE
                )
                try:
                    print(account.root.tree())
                except:
                    pass

if __name__ == "__main__":
    main()
