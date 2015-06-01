Feature: Create DOCX Document
    In order to have a docx file
    As developer
    I want to generate a docx file using Python syntax

    Scenario: Blank document
        Given I have a blank document
        When I save the blank document
        Then DOCX file is created
