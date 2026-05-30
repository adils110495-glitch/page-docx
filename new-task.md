Can you create functionality that automatically generates a separate tab for each language code detected in the slug?

Requirements:

* Parse all language codes from the slug dynamically.
* Create a dedicated tab for every detected language code (e.g., `en`, `fr`, `de`, `es`, etc.).
* Tabs should be generated automatically without hardcoding language codes.
* When a user clicks a language tab, display the document/content associated with that language.
* All language tabs should be contained within the new **Language-Based Document Generator** section.
* The system should support any number of language codes and automatically handle newly added languages.
* The functionality must be completely independent of the existing **DocX Generator** tab while maintaining the same UI/UX style for consistency.
* If no language code is found, display a default tab or an appropriate message.