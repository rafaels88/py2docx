Feature: Create a Paragraph
    In order to have a paragraph
    As developer
    I want to create a paragraph inside the DOCX document

    Scenario: Blank paragraph
        Given I have a docx document instance
        When I get a new paragraph instance
        When I append the new paragraph instance
        Then a blank paragraph is added
