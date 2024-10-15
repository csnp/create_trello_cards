
# **Word Document Formatting Guide**

Properly format your Word document to ensure the script reads it correctly and creates Trello cards based on the provided structure.

---

## **Structure Overview**

To generate Trello cards, your document must follow this structure:

- **Card Title**: Use **Heading 1** style for each card title.
- **Description**: Enter regular text directly below the title.
- **Optional Fields**: Include any of the following, formatted as specified.

---

## **Optional Fields**

Each field is optional but will allow you to customize your Trello cards further. Include any combination of these fields after the card description.

- **Labels**:

  Assign one or more labels to your Trello card. Separate multiple labels with commas.

  `Labels:` label1, label2

- **Due Date**:

  Set a due date for your card using the ISO 8601 format (`YYYY-MM-DDTHH:MM:SS`).
  
  `Due Date:` 2024-12-31T23:59:00

- **Members**:

  Assign one or more Trello members to the card by listing their usernames, separated by commas.
  
  `Members:` username1, username2

- **List**:

  Specify which list on your Trello board the card should be added to.
  
  `List:` List Name

- **Checklist**:

  Add a checklist to your card. Each checklist item should be a bullet point (or numbered list item).
  
  `Checklist:`

  • First item

  • Second item

- **Attachments**:

  Add one or more attachments to your card. These can be URLs or local file paths. Add each attachment on a new line.

  `Attachments:`

  <https://example.com/document.pdf>

  C:\path\to\file.pdf

- **Image** (Cover Image):

  Set a cover image for your card by adding an image URL. The image will be displayed as the card's cover.

  `Image:` <https://example.com/image.png>

---

## **Example Cards**

Here’s a full example of how to format a Word document with two cards, including all optional fields.

### **Sample Card 1: With Full Detailed Card**

# Sample Card Title 1

This is the description of the card.

`Labels:` Marketing, Urgent

`Due Date:` 2024-12-31T23:59:00

`Members:` username1, username2

`List:` To Do

`Checklist:`

• First item

• Second item

`Attachments:`

<https://example.com/document.pdf>

C:\path\to\file.pdf

`Image:` <https://example.com/image.png>

## **Sample Card 2: with Minimal Fields**

# Sample Card Title 2

Here is another card description. This one has fewer optional fields.

`List:` Backlog

`Checklist:`

• Read the docs

• Review the tasks

`Attachments:`

<https://example.com/review.pdf>

---

## **Key Points to Remember**

- **Card Title & Description** Must always be used.
- **Optional Fields** like Labels, Due Date, Members, etc., must follow the exact format as shown above.
- **Multiple Cards**: You can add multiple cards in the same document by repeating the structure (start each card with a Heading 1 title).
- **Checklist Items**: Use bullet points for checklist items.
- **Attachments**: You can include both URLs and local file paths.

---

By following these guidelines, you ensure that the script correctly processes your Word document and creates the corresponding Trello cards.
