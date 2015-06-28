# Php Sharepoint Connector

Back in the days, I created this. If you can use it, you are welcome :)
It was used to connect to a remote Sharepoint 2003 installation, using the (old) webservices.

There are two classes:
* SharepointSite, the first one and handles all calls to the webservices, instantiate with url and credentials
* listQuery for building CAML queries, which can be executed through the listContents method of the SharepointSite class

