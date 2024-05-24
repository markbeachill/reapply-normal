This is a macro for Microsoft Word. 

This was created to format text from ChatGPT to Word (I am  creating materials for students.).
Basically A) paste code from ChatGPT to Dilligner or similar Markdown editor - I like PanWriter
B) Copy preview to Word. (This preview has Normal style applied but it needs reapplying. Reapplying style on WOrd is a "crap shoot".)
C) Apply the formatting here

BUT the code also helps when you have to reapply the style without formatting such as bold and titalics and numbering being removed.

There is a ReapplyNormal function which reapplies normal but preserves: 
- list numbering
- bullet points
- bold and italics at an individual level

There is an ApplyFormattng which:

- Reapplies normal (as above)
- Creates a new table style
- Creates a new table text style
- Applies the new table style to all tables
- Applies the new table text  to all text in tables

Hope this is useful

