# Athletes Administration with Google Apps

<hr>

## Context
In a gym each year there are several athletes that subscribe for the annual training period. To improve the management of these athletes, and to let the obsolete paper system go, a new methodology is required. The fastest and effortless way is to exploit the functionalities that Google suite provide to its users. 

<hr>

## Requirements
- Each athlete must provide a series of personal data, that is: 
    1. Name
    2. Surname
    3. Birth date
    4. Town of birth
    5. City of birth
    6. Residence address
    7. Town of residence
    8. City of residence
    9. Fiscal code
    10. Parent's name
    11. Parent's surname
    12. Contact email
    13. Contact number
    14. Health certificate's expiration date (for people over 6 yo)

    \
    Note that an athlete can't attend the courses without a valid health certificate 

- Each athlete must provide the following documents
    1. Health certificate (for people over 6 yo)
    2. Document for privacy
    3. First payment receipt (between September and December)
    4. Second payment receipt (before the end of January)

<hr>

## Additional features
### Emails
To make things a little more *fancy*, it has been decided to implement some notification sent to all the athletes under specific events. The events are: 
    1. The athlete submits a registration or modify a previous response or someone of the staff modify the registration for him/her. 
    2. The health certificate is going to expire. Athletes between 11 and 18 years old must be notified with a large advance, all the others a couple of months is sufficient.
    2. Gentle reminder for payments  

### Calendar
Since it's important for the staff to check that all the athletes which are doing activities have a valid certificate (possible legal consequences), a notification is sent to the admins. To manage the events, it has been decided to integrate Google calendar and a specific calendar which contains all the expiration dates of the subscribers.  

<hr>

## Execution flow
Since Google Scripts does't let developers to create folder, all the scripts and the email templates are in the same root folder of the project. 

When a new athlete fill the Google form, a new entry is automatically added in the Google Sheets sheet linked to the module. By setting as trigger the answer submit, the `manageAnswers()` function is executed. Here the main steps are performed:
 1. Get the link to modify the response (to be sent later within the email of success registration) from the module
 2. Change the links of the uploaded files (e.g. `https://drive.google...`) in a custom label
 3. Format the text in the other fields (e.g. from `jOhN` to `Jhon`)
 4. Add some formulas like age from birth date and `0`/`1` values on the payment fields whether or not there is a uploaded receipt (i.e. there is the hyperlink pointing to the document) 
 5. Add an event in Google Calendar
 6. Send the email

The result shoud be from this 

| Chronologic info  | Form email | Athlete name | Athlete surname | Birth date | Age | Town of birth | City of birth | Residence address | Town of residence | City of residence | Fiscal code | Contact email | Contact phone number | Health certificate's expiration date | Parent's name | Parent's surname | Health certificate | Privacy document | Receipt 1st payment | Receipt 2nd payment | 1st payment | 2nd payment | Link to modify response |
| ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- |
| 17/09/2024 22.06.10  | athlete@aaga.com  | aJeJe  | braZORf  | 01/01/2000  |    | Lignano  | Venezia  | via Borgo 31  | Lignano  | Venezia  | AAABBB00C00L111D  | athlete@aaga.com  | 1234567890  | 18/09/2024  | gERry  | sCotti  | https://drive.google.com/abc  | https://drive.google.com/def  | https://drive.google.com/ghi  | https://drive.google.com/lmn  |    |    | https://docs.google.com/response |

to this

| Chronologic info  | Form email | Athlete name | Athlete surname | Birth date | Age | Town of birth | City of birth | Residence address | Town of residence | City of residence | Fiscal code | Contact email | Contact phone number | Health certificate's expiration date | Parent's name | Parent's surname | Health certificate | Privacy document | Receipt 1st payment | Receipt 2nd payment | 1st payment | 2nd payment | Link to modify response |
| ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- |
| 17/09/2024 22.06.10  | athlete@aaga.com  | Ajeje  | Brazorf  | 01/01/2000  |  24  | Lignano  | Venezia  | via Borgo 31  | Lignano  | Venezia  | AAABBB00C00L111D  | athlete@aaga.com  | 1234567890  | 18/09/2024  | Gerry  | Scotti  | [Link health certificate](https://drive.google.com/abc)  | [Link privacy document](https://drive.google.com/def)  | [Link receipt 1st payment](https://drive.google.com/ghi)  | [Link receipt 2nd payment](https://drive.google.com/lmn)  | 1 | 1 | [Link to modify response](https://docs.google.com/response) |

 <hr>

 ## Tools
 To create the emails, [Stripo](https://stripo.email/it/) website has been used. 