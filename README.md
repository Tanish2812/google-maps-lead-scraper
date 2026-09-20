# Google Maps Lead Scraper

Scrapes business listings from Google Maps into a clean, formatted Excel leads file.

Built for a marketing agency that needed a leads list of plumbers in Toronto. Produced 115 businesses — names, phone numbers, addresses, websites and ratings.

## What it collects

| Column | Notes |
| --- | --- |
| Business Name | |
| Ad | Yes/No |
| Job | the category Google assigns |
| Address | full street address |
| Phone Number | blank if not listed |
| Rating | stored as a number, so Excel sorts it correctly |
| Reviews | stored as a number |
| Is Open | |
| Has Website | Yes/No |
| Website | blank if none |
| Google Maps Link | |

## Sample output

`sample_output.xlsx` — 20 real rows, formatted exactly like the real deliverable.

Phone numbers are masked (`+1***********2`) so a sample can be shared publicly. The full output is deliberately not included in this repo.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python scraper.py
```

Change `BASE_URL` near the top of the script to search a different trade or city.

## How it works

**1. It scrolls the listing panel on google maps and when the scroll reaches to the end and nothing new loads, it breaks the loop.**

**2. Then it scrapes the required data one by one and stores them in a variable named records.**

**3. Then it stored the recorded data into the Excel file with clean formatted data.**


## Known limitation

Sometime Google Maps don't render the reviews for some listings, hence the script store them as empty. When you see the first few rows are empty, try to rerun the script as it solves automatically when you rerun again. 


## Note on use

This collects publicly visible business information for lead generation. Check that your own use complies with Google's terms and your local data protection rules.

## Built with

Python · Selenium · pandas · openpyxl
