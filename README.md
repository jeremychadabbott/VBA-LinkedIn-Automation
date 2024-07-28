# VBA-LinkedIn-Automation
LinkedIn employers post jobs that say 'remote' but actaully aren't remote - they require travel, or on site.
This code filters against key words to eliminate those fake remote jobs, because I am looking for a truly remote job
this code:
opens my linkedIn profile
clicks Jobs
Clicks 'more'
starts to iterate over the job listings on the left
as it looks at each job description, it refers to a set of 'anti-keyords' such as 'Travel', 'On Site', 'Hybrid' etc.
if a keyword is in the job description, tab 2 is closed and 'X' is clicked on the job so it "wont show it to me anymore"
if no keywords are detected, the program halts so user can view the job. 
the program throws the user a message "hit <enter> to resume" so once user decides whether to apply or not
the program can resume

Anecdote: using this program, out of 300 job listings for me that were "remote", only 10 made it through my filters
saving an enormous amount time.


