I have an existing PHP application that generates multilingual content and currently exports everything into a single document.

I want to add Google Docs integration and automatically create Google Docs Tabs based on language codes extracted from slugs.

Slug examples:

/compensation/ajet/
/fr/compensation/ajet/
/de/compensation/ajet/
/es/compensation/ajet/
/it/compensation/ajet/

Language detection rules:

1. Extract the language code from the slug.

2. If a valid language code exists at the beginning of the slug, use it as the tab name.

3. If no language code exists in the slug, automatically assign `en`.

4. Examples:

   /compensation/ajet/            → en
   /fr/compensation/ajet/         → fr
   /de/compensation/ajet/         → de
   /es/compensation/ajet/         → es
   /it/compensation/ajet/         → it

5. The system must work dynamically for any future language code.

Required functionality:

1. Integrate Google Docs API into the existing PHP application.

2. Create a new Google Document automatically.

3. Extract language codes from generated slugs.

4. Create one Google Docs Tab per language.

5. Use the language code as the tab title.

6. If no language code is found, create/use the `en` tab.

7. Insert the corresponding content into its matching tab.

8. Support unlimited languages without hardcoding.

9. Preserve formatting:

   * Headings
   * Paragraphs
   * Lists
   * Tables
   * Links
   * Bold/italic formatting

10. Keep all languages inside a single Google Document.

11. Generate complete production-ready PHP code.

12. Include:

    * Composer dependencies
    * Google Cloud setup
    * OAuth authentication
    * Google Docs API integration
    * Google Drive API integration
    * Language extraction function
    * Tab creation logic
    * Content insertion logic
    * Error handling
    * Logging

13. Follow PSR standards and OOP architecture.

14. Do not break any existing document-generation functionality.

15. Analyze the current Google Docs API capabilities first and determine whether Google Docs Tabs can be created programmatically.

16. If Google Docs Tabs are not yet fully supported through the API, provide the best available workaround while keeping the code ready for future tab support.

Expected result:

Generated Google Document

Tabs:
├── en
├── fr
├── de
├── es
├── it

Each tab contains only the content associated with that language.
