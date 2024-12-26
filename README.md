# card-Analysis
The provided Python code generates a Word document where each card's details are organized in a structured format. Below is a breakdown of what each part of the code does:

### Code Breakdown:

1. **Data Setup**:
   -  a pandas DataFrame (`df`) from this data.

2. **Word Document Setup**:
   - The `Document()` function from the `docx` module is used to create a new Word document (`doc`).
   
3. **Loop through Each Card**:
   - The `df.groupby('Card Name')` groups the DataFrame by the `Card Name` column, which allows us to process each unique card one by one.
   - For each unique card (`card_name`), the code enters the loop and processes all rows corresponding to that card.

4. **Adding Card Details to the Word Document**:
   - For each card:
     - `doc.add_heading(f"Card Name: {card_name}", level=1)` adds the card name as a heading in the document.
     - The card's category is retrieved from the group and added using `doc.add_paragraph`.
     - A static "Effect" description is added (since it is the same for all levels in this example).
     - A section titled "Level Effects" is added.
     - A second loop (`for index, row in group.iterrows()`) iterates over each level's data for that card and appends the description of each level using `doc.add_paragraph(f"Level {row['Level']}: {row['Description']}")`.

5. **Save the Document**:
   - After processing all cards, the document is saved as `All_Cards_Formatted.docx` using the `doc.save()` method.
   - Finally, a success message is printed: `"Document saved successfully as 'All_Cards_Formatted.docx'"`.

### Output:
The output Word document will be formatted with headings for each card's name and details for each level. The structure looks like this:

#### Example Output:

```
Card Name: Polluted Environment
Card Type: Healthy
Effect: -9 seconds of life per real second in polluted environments.

Level Effects:
Level 1: Lose -4 more seconds per real second in polluted environments.
Level 2: Lose -6 more seconds per real second in polluted environments.
Level 3: Lose -8 more seconds per real second in polluted environments.
Level 4: Lose -10 more seconds per real second in polluted environments.
Level 5: Lose -12 more seconds per real second in polluted environments.
Level 6: Lose -9 more seconds per real second in polluted environments.
Level 7: Lose -10 more seconds per real second in polluted environments.
Level 8: Lose -11 more seconds per real second in polluted environments.
```

Each card is processed in sequence and added to the Word document with its name, type, general effect, and detailed level effects. The document is saved as `All_Cards_Formatted.docx`.

### Key Libraries:
- **pandas**: Used to handle the tabular data (DataFrame).
- **python-docx**: Used to generate and manipulate the Word document.

---

### Additional Notes:
- Replace the sample data (`data` dictionary) with your actual DataFrame.
- You can modify the `Effect` section to include dynamic values if needed.
- The document's structure and formatting can be further customized as per your requirement.
