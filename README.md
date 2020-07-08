# GetPhoneList
A PowerShell script that queries Active Directory for enabled users, sorts them by department and job title, exports the list to csv, then copies the csv to a folder

# Description
Let say you have a company intranet page with everyones contact information. Let's say you store said contact information in AD. Let's also say that you don't use SharePoint or any other contact management software. Well, this is why this script exists. It's a simple way to extract contact info from AD, sort it how you want, then export it to a csv and save it to a server to be served up on a website.

# Custom Sort
This is one way to customize the order of how names show up. I wanted the President to be first on the last, then VP, Director, etc. This actually took some time for me to figure out. Eventually I got help from the nice people over at StackOverflow. Thanks to Ansgar Wiechers on SO for helping me find a suitable solution to my needs. https://stackoverflow.com/questions/58211308/how-to-sort-and-group-csv-file-containing-a-list-of-employees-by-department-and

```
$sortedEmployees = $searchQuery | 
                 Sort-Object {$departments.IndexOf($_.department)}, 
                 @{expression={$_.Title -match "(President)|(Controller)"}; descending=$true}, 
                 @{expression={$_.Title -match "(VP)"}; descending=$true}, 
                 @{expression={$_.Title -match "(Director)"}; descending=$true},
                 @{expression={$_.Title -match "(Manager)"}; descending=$true}, 
                 @{expression={$_.Title -match "(Supervisor)"}; descending=$true}, Name
```