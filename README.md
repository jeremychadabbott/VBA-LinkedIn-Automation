LinkedIn often lists jobs as "remote" that aren't truly remote—they may require travel or onsite presence. This code filters out these misleading job postings by using specific keywords to identify and eliminate them. The goal is to find genuinely remote job opportunities efficiently.

### How It Works:

1. **Profile Access**: Opens your LinkedIn profile (option 1). option 2 is to search existing profile.
2. **Navigation**: Clicks on the 'Jobs' tab.
3. **Job Listings**: Clicks 'More' to load additional job listings.
4. **Iteration**: Iterates over each job listing on the left panel.
5. **Keyword Filtering**: Examines each job description for a set of 'anti-keywords' such as 'Travel', 'On Site', 'Hybrid', etc.
6. **Action on Keywords**: If an anti-keyword is found, the code closes the job tab and clicks 'X' to ensure the job is not shown again.
7. **User Review**: If no keywords are detected, the program pauses to allow the user to review the job.
    - A message "Hit any key to resume" is displayed, allowing the user to decide whether to apply or not before resuming the filtering process.

### Important Notes:

- **Automatic Mouse Movement**: This code uses automatic mouse movement and targeting instead of an API.
- **Screen Resolution**: The screen resolution must be set correctly for the mouse movements to align with the intended targets.
- **Display Changes**: Any changes in how LinkedIn is displayed may break the code, requiring adjustments to the script.

### Anecdote:

Using this program, I filtered through 300 job listings labeled as "remote." Only 10 passed my filters, saving me an enormous amount of time.

## Instructions
open the XLSM (macro enabled spreadsheet)
there are two buttons. One "Normal run, launch chrome, search jobs" and the other "Normal run on chrome already open"
if you click "Normal run, launch chrome, search jobs"
    chrome will launch and resize.
    chrome will navigate to "linkedin" (there is an assumption here that you are already logged into linkedin"
    using mouse coordinates, mouse clicks on "Jobs" 
    NOTE* mouse coordinate code requires your screen be set to 1366x768
    using mouse coordinates, code clicks on "More"
    now a loop begins:
    using mouse coordinates, mouse click and opens HTML window
    using mouse coordinates, mouse clicks into HTML windo and right clicks "copy" and cliks "outer HTML"
    code parses and finds first job that is NOT ALSO "we won't show you this job aymore" (important later)
    code opens job in new tab
    code copies text of wopen window
    code checks text against KEYWORDS (spreadsheet! modify keywords to suite YOU!)
    if any keywords are found, this is a job we DONT WANT
    if no keywords are found, code halts and waits for user interaction
        if no keywords are found, code presents user "hit <enter> to resume"
        when you hit enter, the screen will flip back to chrome, 2nd tab will close, page will refresh, and serach begins again
    if keywords are found, 2nd tab closes, page text is copied to check where mouse should click to hit "X" (don't show me this job again)
    page refreshes
    search resumes until you hit <break>

    if "Normal run on chrome already open" is clicked
    this assumes you have already opened linkin page and have performed a search and want to use this code against your search results
    code behaves almost the same as above except chrome does not launch so
    first you should have opened linkein. opened jobs. performed a search. now hit "Normal run on chrome already open" button

    This code is just a little hack that helps me screen linkin positions. I leave it running on a second screen and when it halts
    I see if it's a job I want or if I need to add more keywords to further filter out dumb jobs. 

    

    

