
# Redmine Time Registration
Application to automatically register on Redmine your time entries based on a excel sheet.

## Usage
1. Download the executable and the Time Sheet [here](https://1drv.ms/f/s!AryB5hpD8sR2l5EUkkp-WRfiDnX4ZQ).
2. Create a File named "Redmine Time Registration" in yours Document File.
3. Put the two files downloaded in the folder created above.
4. Open the Excel Sheet. 
5. Go to tab _Parameters_. There you have to put some obligatory information.
6. ![Parameter Fields](https://raw.githubusercontent.com/eleonardoro/redmine-time-registration/master/images/Screenshot_4.png)
7. _**Activty Type**_: table where where you put the types of activities that you have in your organization and their ids. To know the all the types and the ids you should go to _Administration_ page. Find _Enumerations_ option. Then, find _Activities (Time Register)_. Now this is a little s**t, you have to open all the itens and look in the URL the id there.
8. ![Path to administration and Enumeration pages](https://raw.githubusercontent.com/eleonardoro/redmine-time-registration/master/images/Screenshot_1.png)
9. ![Actvities Types and an Activity Type](https://raw.githubusercontent.com/eleonardoro/redmine-time-registration/master/images/Screenshot_2.png)
10. ![Activity Page and its ID](https://raw.githubusercontent.com/eleonardoro/redmine-time-registration/master/images/Screenshot_3.png)
11. _**API Acess Key**_:  cell where you put your custom api acess key. To get or generate your custom API you go in _My Account_, the click on show button as shown bellow.**If this option is unnabled, or gives you error that means that your redmine rest api is not configured yet - there is a lot of tutorials on the web to teach how to do it**
12. ![Path to get your API Key](https://raw.githubusercontent.com/eleonardoro/redmine-time-registration/master/images/Screenshot_5.png) 
13. _**URL Redmine**_:  here you put the redmine url (daar).
14. **Done!**  Just double click the executable downloaded and it will do all the hard work!
**Obs.:** In order to use ir right I strongly recomend you to read the _Rules and FAQ_ section right bellow.
**Obs. 2:** This is a beta version, I also strongly recoment to you run the exe and certify in redmine that it is all right.

## Rules and FAQ

 1. **Obligatory Fields**: Date, Hours, Issue and Activity Type.
 2. **Status Field**: Is the field used to let you know the result. If it is OK the program will not even try to register it again, and, anything different than that it will try to register again.
 3. **If openned, close the sheet**: I know it su**s, in the next versions it'll not be necessary, I promisse.
 4. **Time doesn't update**: just add another line in the sheet anywhere and it will update - not my fault
