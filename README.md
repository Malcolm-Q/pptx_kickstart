# pptx_kickstart
A simple script to generate a PowerPoint given a prompt using openai and python-pptx.

# Overview
* [Code](/powerpoint_kickstart.py)
* [First Example](/why%20exercise%20is%20important.pptx)
* [Second Example](/baseball%20in%20Japan.pptx)

# Description
Module I made for my friend to kickstart presentations on general topics. Produces a structured .pptx file with slide topics and context / ideas to talk about. Needs to be formatted further.

# Requirements
* Packages:
    * openai
    * python-pptx
* API Keys:
    * openai API key environment variable set to 'openai_key'

# Notes
You can change the system prompt as you like, the api request has not been fine tuned to save API cost. Because of this it will sometimes generate powerpoints that have structure but a lack of notes.