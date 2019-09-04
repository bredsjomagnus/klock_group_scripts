# klock_group_scripts

klock_grupper reads csv downloaded from Infomentor with students name and
groups. Then create one csv-file for every group that can be used to
create these groups in Google Admin Groups.

To make this work there needs to be a reference csv with the students
names and email addresses [elevnamn_till_elevmail.csv]

USAGE
$ python klock_grupper.py csv-file
The csv-file needs to have semicolon as separator. But creates csv-files with comma as separator.

csv-file head example from Infomentor:
Elev Grupper;Elev Namn;Elev Klass;Årskurs
9ABCNO-2, 9CBL-1, 9CEN, 9CHKK-1, 9CIDH, 9CMU-1, 9CSL-1, 9CSO;last name, first name;9C;9
...

Then creates files like this:
Group Email [Required],Member Email,Member Type,Member Role
groupmail@edu...,firstname.lastname@edu.hellefors.se,USER,MEMBER
...

Options:
    -h or --help        Display this help message