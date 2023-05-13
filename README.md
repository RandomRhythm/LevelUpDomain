### Level Up Domain
#### Takes a list of domain names and creates a sampling list by only outputting one subdomain for each primary domain.

This script will parse a provided list of domain names to perform the following:
* Remove invalid domains
* Separate IP addresses from the domains
* Output primary domain name prevalence 
* Output the list of primary domain names
* Output a domain sampling 

Domain names have a top-level domain (TLD) and then subdomains. However, not all subdomains immediately beneath a TLD are unique to an owner/operator. Such a subdomain is known as a [public suffix](https://publicsuffix.org), which is a domain that allows for domain registration beneath it. This script utilizes lists from [surbl.org](http://www.surbl.org/news/internal/Added-domains-to-two-level-tlds-and-three-level-tlds), [iana](http://data.iana.org/TLD/tlds-alpha-by-domain.txt), and internal checks to determine the highest level of a domain structure that is unique to an owner. In other words, subdomains are removed to identify the primary domain name that a registrar would sell for registration. After acquiring the primary domain name, then a sampling of full domain structures can be obtained. 

The Level Up Domain script will create the following files:
* Domain_Sample.txt - Sample of domain names.
* DomainPrev.csv - Prevalence based on parent domain.
* LevelUp_Domains.txt - List of primary domain names.
* IP_Addresses - IP addresses identified in the list of domains.
* Invalid_Domain_IP.txt - List of items provided that were an invalid IP address or domain name.

Warning! The input file only supports ANSI and UTF-16 LE encoding. Use notepad Save As dialog or other method to change the encoding beforehand.
