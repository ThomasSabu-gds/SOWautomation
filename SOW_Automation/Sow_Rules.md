Define the rules to generate the final Word document using:
Excel input (SOW Text column)
User Yes/No selection
Template highlighted text and NOTE TO DRAFT rules
RULES

**2.1 NOTE TO DRAFT**
&nbsp; 1.All yellow highlighted text needs to be either updated or deleted
&nbsp; 2.All text enclosed in \[ ] need to be either updated or deleted
&nbsp; 3.All text enclosed in \[ ] and starting with "NOTE TO DRAFT" need to be deleted
&nbsp; 4.All text enclosed in \[ ] and not starting with "NOTE TO DRAFT" need to be updated or deleted
&nbsp; 5.All text highlighted in blue need not be touched unless specifically mentioned in the clause wise rules

**\[IMPORTANT]**
&nbsp; Remove all text marked as “NOTE TO DRAFT” inside \[ ].
&nbsp; If it appears within a paragraph, remove only that part.
&nbsp; Do not delete the full paragraph.
&nbsp; Keep the remaining sentence grammatically correct.

**2.2 Highlighted Text\[approach taken is match first and update]**
&nbsp; If matching data exists in the Excel SOW Text column → Replace the highlighted text.
&nbsp; If no matching data exists → Delete the entire highlighted section.

**2.3 User Selection Logic**
&nbsp; If user selects YES:
&nbsp; Replace highlighted text with the value entered in the UI.
&nbsp; If the section contains “NOTE TO DRAFT”, remove that part before replacing.
&nbsp; If user selects NO:
&nbsp; Delete the entire related highlighted section.

**5. Schedule Deletion Rules\[ CURRENT APPROACH,STILL WORKING TO DECIDE THE APPROCH]**
&nbsp; 1.Clause numbers beginning with “Sch” shall represent an entire Schedule.
&nbsp; 2.Where the User selects “NO” for a Schedule clause, the corresponding Schedule shall be deleted in its entirety.
&nbsp; 3.A Schedule shall commence at a paragraph formatted as Heading 1 with the text:
&nbsp; Schedule <Letter>
&nbsp; 4.Deletion shall continue up to, but not including, the next paragraph formatted as Heading 1 with the text:
&nbsp; Schedule <Letter
&nbsp; 5.Any other Heading 1 paragraphs appearing within the Schedule shall be disregarded for boundary determination.
&nbsp; All content within the identified Schedule boundaries shall be deleted, including:
