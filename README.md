This is a Microsoft Word macro that reaplies the Normal style while maintaining bold, italics etc. 
Word does not reapply styles such as normal very well - it has a strange view of when to ignore or use bold/italics.

The ReapplyNormal macro/function reapplies the Normal style formatting but preserves: 
- list numbering
- bullet points
- bold and italics at an individual word level

The ApplyFormattng macro/function:
- Reapplies normal (as above)
- Creates a new table style
- Creates a new table text style
- Applies the new table style to all tables
- Applies the new table text  to all text in tables

It can be useful for:
- fixing pasted Mardown text converted to html (from Dillinger, PanWriter etc)
- fixing where documents dont use styles consistently

It includes some optimisation - but this is still not an instant macro - so a bit of patience is needed.

Hope this is useful - let me know if it is.

This is  public domain.

This was created to format text from ChatGPT to Word (I am  creating materials for students).
Its core is updating the formatting to match the styles for Normal, i.e. reapplying the Normal style


