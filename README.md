# One Note to Markdown Converter

Convert One Note files to the markdown format

## Obtaining your access token

1. Go to [graph explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in with your Microsoft account.
2. Click on the "Access token" tab copy the token inside.

## How to Use

1. Clone this repository
2. Install dependencies (requests and beautifulsoup)
3. Run the script ```python app.py```
4. Follow the instructions. Input your access token. Select whether you want to prefix the notes with dates. If so, select a date format. Select a notebook to convert.
5. Wait for the script to finish.
6. Find the converted markdown files in the `onenote_ouput` directory.

## Clearing Cache

- Cache is stored in the `onenote_output` directory in the `.conversion_cache.json` file.
- This file tracks which files have been converted, preventing duplicate conversions. The main reason for this is if the script fails and you need to run it a second time, you won't have duplicates and you won't be making as many server requests.
- To clear the cache, you can either run the `clear_cache.py` script, or simply delete the conversion cache file.

## Why I made this

- I desperately wanted to leave One Note and the Microsoft ecosystem which I have felt locked into for many years. I wanted to switch my older files to Markdown for a while, especially since moving to Linux.
- There are other solutions out there, but I wanted one that didn't involve other software and was open-source. Additionally, most other converters involved some kind of Microsoft involvement, whether it was made for Windows or involved the `.docx` format in some way.
- This program should address the need for a one-note to markdown converter, at least for me.

- Disclaimer 1: I'm a software dev but not a python dev, so this may be flaky and syntactically messed up.
- Disclaimer 2: I did in fact use AI to help me write this program, due to disclaimer 1.

