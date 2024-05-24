This is a macro for Microsoft Word. 

This was created to format text from ChatGPT to Word (I am  creating materials for students.).
Its core is updating the formatting to match the styles for Normal, i.e. reapplying the Normal style

Basically 
A) paste the copied text from ChatGPT into Dilligner or similar Markdown editor - I like PanWriter
B) Paste preview from editor into Word.
C) Run the macro in Word to make it look nice

The code also helps when having to fix a document where you need to reapply the style and preserve formatting such as bold and italics and numbering.

There is a ReapplyNormal function which reapplies the Normal style formatting but preserves: 
- list numbering
- bullet points
- bold and italics at an individual word level

There is an ApplyFormattng which:

- Reapplies normal (as above)
- Creates a new table style
- Creates a new table text style
- Applies the new table style to all tables
- Applies the new table text  to all text in tables

It includes some optimisation - but this is still not an instant macro - so a bit of patience needed.

Hope this is useful - let me know if it is.

This is  public domain.
