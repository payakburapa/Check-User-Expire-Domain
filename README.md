# Job Check update expire Date Username login windows

OLD JOB CHECK BY MANUAL

1. Admin check User expire by command line "net user XXXXXXX /DOMAIN"

2. Command will show Detail expire and update last

3. But user have many

-----------------------------------------

NEW JOB CHECK BY PYTHON PROGRAM 

1. Python run command line "net user XXXXXXX /DOMAIN > user.txt" by get user from excel column 'User Owner'

2. Program read file text user.txt

3. Program substring online date account expire

4. Program check by date expire subtract date now

5. Program get result from step 4 input to excel column 'Remain day' row user login

6. Program check result. if Less than 8 will change color word to red

7. Program will loop since step 1 to step 6 to finished