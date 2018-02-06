
# Redmine Time Registration
Application to automatically register on Redmine your time entries based on a excel sheet.

## Usage
1. Download the executable and the Time Sheet [here](https://1drv.ms/f/s!AryB5hpD8sR2l5EUkkp-WRfiDnX4ZQ).
2. Create a File named "Redmine Time Registration" in yours Document File.
3. Put the two files downloaded in the folder created above.
4. Open the Excel Sheet. 
5. Go to tab _Parameters_. There you have to put some obligatory information.
6. ![Parameter Fields](https://xymu7a.dm2301.livefilestore.com/y4pBf1DNON3NqLPnQptMXodQp1rf6U28iZLjgAEHYZE7HMjkmDERP8PHyou0jaxzy9_TmLpq_7NeeZdxkdEunThHEw21gFKOTjCQ3a3JaU8NAJSsZzSDsP2CcavY5TY50l_SFoiqNpzWrjkIasP0ddcx82hR9dGDdFaEh9hLNBxXsP8E1lefk_29pDeYO6-DC6vGUdWx9GV-0_Ciw014UnUSw/Screenshot_4.png?psid=1)
7. _**Activty Type**_: table where where you put the types of activities that you have in your organization and their ids. To know the all the types and the ids you should go to _Administration_ page. Find _Enumerations_ option. Then, find _Activities (Time Register)_. Now this is a little s**t, you have to open all the itens and look in the URL the id there.
8. ![Path to administration and Enumeration pages](https://xymu7a.dm2301.livefilestore.com/y4pV29I5Kb03Pn5zxPCIAYVINq_DSLJfSKxdjBQz5fu4q8ohca1wxGRLmHbe-WtuxQxixGAbAgIXsyKLeUYk5fmEnC09gTSDUdAkASA4EUc7Nx7gPmVyJp9r_3lCBr3FUb8aDXhIoWmY99nBQTjfdgf-D2tJ5nUg3hyb4QrtjR7_yYj9ybPywosoYfSVQIX90Eq1MBDX328DzdVCDVtyJYnEg/Screenshot_1.png?psid=1)
9. ![Actvities Types and an Activity Type](https://xymu7a.dm2301.livefilestore.com/y4pPIKFwsc4uoPJWgNuYO_p9JATti418-tralZ6J78hpW1yR6Ko38SzVtisXIp62sq34osWDinKX_o3i4DKyJzJtQN-H8eUmxy7dNECm-71firpphCkry3Y6j5JnZvvf5_OfPKOx1ClPTTTGlU-Kn-0OFMkvl0agAZufTw72aPIwpFZN_A5LEolOiNf94t1o95hWdtCG58-ittODvfXXfBncQ/Screenshot_2.png?psid=1)
10. ![Activity Page and its ID](https://xymu7a.dm2301.livefilestore.com/y4pp4vwN3FLI_tOGtcr3FMoohiYnbLcQYSsrv2CCTGTC8tdmV1YKAg4KpNjRi_PhHA76tQxYuE9KxBFfxrdckQpq9pXmS3l4dwzGdEhd0x2576i46BXNmhXN4IDmjlclt6R8dfvIG_gxQLrm1h3yZz4aqn4kUXkUpAhedE2t6X_JhSg1NY_6TDygtm7wt4NWbgm2VfgD5ZeR3e8kK15zIbaqw/Screenshot_3.png?psid=1)
11. _**API Acess Key**_:  cell where you put your custom api acess key. To get or generate your custom API you go in _My Account_, the click on show button as shown bellow.**If this option is unnabled, or gives you error that means that your redmine rest api is not configured yet - there is a lot of tutorials on the web to teach how to do it**
12. ![Path to get your API Key](https://xymu7a.dm2301.livefilestore.com/y4p91O0pcXytZsJddg7XkzuDTYd1CoLl1XhZZp4kICEjYuyygVpdrw53XnAnlVlRgS2vhavHw5mRwbPansCn6mjdnNj27lEL9YO6drWIU5iVP8x4v6FxPrOX2lbdQbgsbrbv3PmHM_UutdWrhaC9qJ_xRAjkkHvWnfDoGW7R5KUwtBaVSds1CHWzJRdOpciNR-aInZ2OD8NN-KiWRblQBzg4g/Screenshot_5.png?psid=1) 
13. _**URL Redmine**_:  here you put the redmine url (daar).
14. **Done!**  Just double click the executable downloaded and it will do all the hard work!
**Obs.:** In order to use ir right I strongly recomend you to read the _Rules and FAQ_ section right bellow.
**Obs. 2:** This is a beta version, I also strongly recoment to you run the exe and certify in redmine that it is all right.

## Rules and FAQ

 1. **Obligatory Fields**: Date, Hours, Issue and Activity Type.
 2. **Status Field**: Is the field used to let you know the result. If it is OK the program will not even try to register it again, and, anything different than that it will try to register again.
 3. **If openned, close the sheet**: I know it su**s, in the next versions it'll not be necessary, I promisse.
 4. **Time doesn't update**: just add another line in the sheet anywhere and it will update - not my fault
