// --- Configuratie ---

var ADDON_VERSION = "0.11"; // <<< Versienummer van de Add-on
var DEFAULT_LABEL_NAME = "---"; // Standaard label als niets is ingesteld
var TARGET_SHEET_NAME = "Gmail DocLog Sheet"; // De standaard/gewenste naam voor de sheet
var TARGET_TAB_NAME = "Data"; // De naam van het tabblad in de sheet
var SHEET_ID_PROPERTY_KEY = 'docLogSheetId'; // Sleutel voor UserProperties om de sheet ID op te slaan
var LABEL_NAME_PROPERTY_KEY = 'docLogLabelName'; // Sleutel voor UserProperties om de labelnaam op te slaan


// --- Hoofd Entry Point (Homepage Trigger) ---

/**
 * Bouwt de initiële kaart. Controleert sheet en daarna het label.
 * @param {Object} e Event object van de trigger.
 * @return {Card[]} Een array met de kaart die getoond moet worden.
 */
function buildHomepageCard(e) {
  var userProperties = PropertiesService.getUserProperties();
  var sheetId = userProperties.getProperty(SHEET_ID_PROPERTY_KEY);
  var card;

  // 1. Check/Bepaal Sheet ID
  if (sheetId) {
    try {
      SpreadsheetApp.openById(sheetId);
      Logger.log("Geldige sheet ID gevonden: " + sheetId);
      // Ga verder naar stap 2 (Label Check) hieronder
    } catch (err) {
      Logger.log("Opgeslagen sheet ID '" + sheetId + "' ongeldig/niet toegankelijk: " + err);
      userProperties.deleteProperty(SHEET_ID_PROPERTY_KEY);
      sheetId = null;
      // Ga verder naar de findOrCreateOrPromptSheet logica hieronder
    }
  }

  // Als er geen geldige ID is (of was), zoek/creëer/prompt voor sheet
  if (!sheetId) {
    var findResult = findOrCreateOrPromptSheet();
    switch (findResult.status) {
      case 'found':
      case 'created':
        userProperties.setProperty(SHEET_ID_PROPERTY_KEY, findResult.id);
        sheetId = findResult.id; // Belangrijk: zet sheetId voor de volgende stap
        // Ga verder naar stap 2 (Label Check) hieronder
        break; // Ga NIET meteen de kaart bouwen
      case 'duplicate':
        card = createSelectionCard(findResult.message);
        return [card.build()]; // Toon sheet selectie kaart en stop hier
      case 'error':
      default:
        card = createErrorCard(findResult.message);
        return [card.build()]; // Toon foutkaart en stop hier
    }
  }

  // 2. Check Label (alleen als we een geldige sheetId hebben)
  if (sheetId) {
      var configuredLabelName = getConfiguredLabelName();
      // Controleer of het label bestaat, maar maak het NIET aan bij alleen opstarten
      var labelExists = checkAndEnsureLabelExists(configuredLabelName, false);

      if (labelExists) {
          // Label bestaat, toon de hoofdactiekaart
          Logger.log("Geconfigureerd label '" + configuredLabelName + "' gevonden. Hoofdkaart wordt getoond.");
          card = createMainActionCard(sheetId);
      } else {
          // Label bestaat niet, toon de label setup kaart
          Logger.log("Geconfigureerd label '" + configuredLabelName + "' NIET gevonden. Label setup kaart wordt getoond.");
          card = createLabelSetupCard(configuredLabelName); // Geef de niet-gevonden naam mee
      }
  } else {
      // Dit zou niet mogen gebeuren door de logica hierboven, maar als fallback
      Logger.log("Fout: Kon geen sheet ID bepalen voor label check.");
      card = createErrorCard("Kon de spreadsheet niet instellen. Probeer opnieuw.");
  }

  return [card.build()]; // Geef de gebouwde kaart terug
}


/** --- functie om de opgeslagen labelnaam op te halen
 * Haalt de door de gebruiker geconfigureerde labelnaam op uit UserProperties.
 * Geeft de DEFAULT_LABEL_NAME terug als er nog niets is ingesteld.
 * @return {string} De geconfigureerde of standaard labelnaam.
 */
function getConfiguredLabelName() {
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(LABEL_NAME_PROPERTY_KEY) || DEFAULT_LABEL_NAME;
}


// --- Functie om Sheet te Vinden, Maken of Selectie te Vragen ---
/**
 * Zoekt naar de sheet op naam, maakt deze aan indien niet gevonden,
 * of geeft status 'duplicate' terug als er meerdere zijn.
 * @return {object} Object met { status: 'found'|'created'|'duplicate'|'error', id: sheetId | null, message: errorMsg | null }
 */
function findOrCreateOrPromptSheet() {
  try {
    var files = DriveApp.getFilesByName(TARGET_SHEET_NAME); // Zoek naar bestanden met de doelnaam
    var foundFiles = [];
    while (files.hasNext()) {
      foundFiles.push(files.next()); // Verzamel alle gevonden bestanden
    }

    if (foundFiles.length === 1) {
      // Precies één gevonden! Perfect.
      var fileId = foundFiles[0].getId();
      Logger.log("Precies één sheet gevonden op naam: " + TARGET_SHEET_NAME + " (ID: " + fileId + ")");
      return { status: 'found', id: fileId, message: null };

    } else if (foundFiles.length === 0) {
      // Geen gevonden! Maak een nieuwe aan.
      Logger.log("Sheet '" + TARGET_SHEET_NAME + "' niet gevonden, wordt aangemaakt...");
      var ss = SpreadsheetApp.create(TARGET_SHEET_NAME); // Maak nieuwe sheet aan in root Drive
      var newSheetId = ss.getId();
      Logger.log("Sheet succesvol aangemaakt met ID: " + newSheetId);

      // Voeg meteen het Data tabblad en de juiste headers toe
      var sheet = ss.insertSheet(TARGET_TAB_NAME);
      sheet.appendRow(["Datum", "Afzender", "Ontvanger", "Titel", "Beschrijving", "Email Link", "Attachment/Link"]); // Headers in juiste volgorde

      // --- LOGICA OM STANDAARD BLAD TE VERWIJDEREN ---
      var allSheets = ss.getSheets(); // Haal alle bladen op (zowel het originele als "Data")
      for (var s = 0; s < allSheets.length; s++) {
          var currentSheetName = allSheets[s].getName();
          if (currentSheetName !== TARGET_TAB_NAME) { // Als de naam NIET "Data" is...
              try {
                  ss.deleteSheet(allSheets[s]); // ...verwijder het dan.
                  Logger.log("Standaardblad '" + currentSheetName + "' verwijderd.");
              } catch (delErr) {
                  Logger.log("Kon blad '" + currentSheetName + "' niet verwijderen: " + delErr);
              }
          }
      }
      // --- EINDE LOGICA OM STANDAARD BLAD TE VERWIJDEREN ---

      return { status: 'created', id: newSheetId, message: "Nieuwe sheet '" + TARGET_SHEET_NAME + "' is aangemaakt." };

    } else {
      // Meerdere gevonden! Dit vereist handmatige selectie door de gebruiker.
      Logger.log("Fout: Meerdere spreadsheets gevonden met de naam '" + TARGET_SHEET_NAME + "'. Aantal: " + foundFiles.length);
      // Maak een lijst van de gevonden bestanden voor de boodschap
      var fileNames = foundFiles.map(function(f){ return "- " + f.getName() + " (ID: ..." + f.getId().slice(-6) + ")"; }).join('\n');
      return {
        status: 'duplicate',
        id: null,
        message: "Er zijn " + foundFiles.length + " spreadsheets gevonden met de naam '" + TARGET_SHEET_NAME + "':\n" + fileNames +
                 "\n\nOpen Google Drive, zoek de juiste sheet, kopieer de ID of URL en plak deze hieronder om door te gaan."
      };
    }
  } catch (e) {
    // Vang eventuele fouten tijdens Drive/Spreadsheet operaties op
    Logger.log("Fout tijdens zoeken/maken van sheet: " + e + "\nStack: " + e.stack);
    return { status: 'error', id: null, message: "Fout bij toegang tot Google Drive/Spreadsheets: " + e.message };
  }
}


// --- Kaart Bouw Functies ---

/**
 * Maakt de hoofdactiekaart die wordt getoond als de sheet bekend is.
 * Bevat knoppen voor verwerken, sheet openen, instellingen (label, reset), en toont versie.
 * @param {string} sheetId De ID van de te gebruiken spreadsheet.
 * @return {CardBuilder} De gebouwde kaart (nog zonder .build()).
 */
function createMainActionCard(sheetId) {
  var card = CardService.newCardBuilder();
  var projectName = ScriptApp.getProjectName(); 

  card.setHeader(CardService.newCardHeader().setTitle(projectName || 'E-mails Verwerken'));  // Pas de titel van de kaart aan om de projectnaam te bevatten (optioneel, maar netjes)

  var currentLabelName = getConfiguredLabelName(); // Haal huidige labelnaam op

  // Sectie met info over de geselecteerde sheet en knop om te openen
  var infoSection = CardService.newCardSection();
  var sheetName = "Onbekend";
  var sheetUrl = "#";
  try {
      var ss = SpreadsheetApp.openById(sheetId);
      sheetName = ss.getName();
      sheetUrl = ss.getUrl();
  } catch(e) {
      Logger.log("Kon naam/URL niet ophalen voor ID " + sheetId + " tijdens kaartbouw: " + e);
      sheetUrl = "https://docs.google.com/spreadsheets/d/" + sheetId + "/edit";
  }

  infoSection.addWidget(CardService.newTextParagraph()
      .setText("Klaar om e-mails met label '<b>" + currentLabelName +
               "</b>' te verwerken.\nDoel Sheet: <b>" + sheetName + "</b>"));

  infoSection.addWidget(CardService.newTextButton()
      .setText("Open '" + sheetName + "'")
      .setOpenLink(CardService.newOpenLink().setUrl(sheetUrl)));

  card.addSection(infoSection);

  // Sectie voor Instellingen
  var settingsSection = CardService.newCardSection().setHeader("Instellingen");

  // Input voor Labelnaam
  settingsSection.addWidget(CardService.newTextInput()
      .setFieldName('label_name_input')
      .setTitle("Te Verwerken Gmail Label")
      .setValue(currentLabelName));

  // Knop om Label op te slaan
  settingsSection.addWidget(CardService.newTextButton()
      .setText("Label Opslaan")
      .setOnClickAction(CardService.newAction().setFunctionName("handleSaveLabelAction")));

  // Reset Knop
  settingsSection.addWidget(CardService.newDecoratedText()
        .setText("Reset Sheet Selectie")
        .setBottomLabel("Vergeet de huidige spreadsheet koppeling.")
        .setButton(CardService.newTextButton()
            .setText("Reset")
            .setOnClickAction(CardService.newAction().setFunctionName("handleClearSheetIdAction"))
            .setTextButtonStyle(CardService.TextButtonStyle.TEXT)));

  // Toon Versienummer 
  // settingsSection.addWidget(CardService.newTextParagraph()
  // .setText("<i>Add-on Versie: " + ADDON_VERSION + "</i>")); 

  // Toon de projectnaam ipv versienummer 
  settingsSection.addWidget(CardService.newTextParagraph()
      .setText("<i>Project: " + (projectName || "Naamloos Project") + "</i>")); 

  card.addSection(settingsSection);

  // Footer met de hoofdactieknop
  card.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Verwerk Nu')
      .setOnClickAction(CardService.newAction().setFunctionName('runProcessingLogic'))));

  return card;
}


/**
 * Maakt een kaart die wordt getoond als het geconfigureerde label niet bestaat.
 * Vraagt de gebruiker om een labelnaam in te voeren (bestaand of nieuw).
 * @param {string} suggestedLabelName De labelnaam die niet gevonden werd (suggestie).
 * @return {CardBuilder} De gebouwde kaart (nog zonder .build()).
 */
function createLabelSetupCard(suggestedLabelName) {
    var card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader().setTitle('Gmail Label Instellen'));

    var section = CardService.newCardSection();
    section.addWidget(CardService.newTextParagraph()
        .setText("Het ingestelde Gmail label '<b>" + suggestedLabelName + "</b>' is niet gevonden." +
                 "\n\nVoer hieronder de naam in van een <b>bestaand label</b> dat u wilt gebruiken, " +
                 "of voer een <b>nieuwe naam</b> in om een label aan te maken."));

    // Input veld voor de gebruiker
    section.addWidget(CardService.newTextInput()
        .setFieldName('label_name_input') // Zelfde naam als op hoofdkaart
        .setTitle('Gmail Label Naam')
        .setValue(suggestedLabelName)); // Vul de niet-gevonden naam alvast in

    card.addSection(section);

    // Knop om op te slaan/aan te maken
    card.setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
            .setText('Gebruik / Maak Dit Label')
            // Gebruikt dezelfde save functie, die nu check/create logica bevat
            .setOnClickAction(CardService.newAction().setFunctionName('handleSaveLabelAction'))));

    return card;
}


/**
 * Maakt de kaart die wordt getoond als meerdere sheets met zelfde naam zijn gevonden.
 * Vraagt de gebruiker om handmatig de juiste sheet te selecteren via ID of URL.
 * @param {string} message De boodschap die de situatie uitlegt.
 * @return {CardBuilder} De gebouwde kaart (nog zonder .build()).
 */
function createSelectionCard(message) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Selecteer Spreadsheet Handmatig'));

  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(message)); // Toon de fout/instructie

  // Input veld voor de gebruiker om ID/URL te plakken
  section.addWidget(CardService.newTextInput()
    .setFieldName('sheet_input') // Naam om de invoer uit te lezen in de actie
    .setTitle('Plak hier de Sheet ID of Volledige URL'));

  // Knop om Google Drive te openen in een nieuw tabblad als hulp
  section.addWidget(CardService.newTextButton()
     .setText('Open Google Drive (nieuw tabblad)')
     .setOpenLink(CardService.newOpenLink().setUrl('https://drive.google.com/drive/u/0/my-drive'))); // Link naar Drive root

  card.addSection(section);

  // Knop om de selectie op te slaan
  card.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Gebruik Deze Sheet')
      .setOnClickAction(CardService.newAction().setFunctionName('saveSelectedSheet')))); // Roept de opslagfunctie aan

  return card;
}

/**
 * Maakt een generieke foutkaart voor onverwachte problemen.
 * @param {string} message De foutboodschap om te tonen.
 * @return {CardBuilder} De gebouwde kaart (nog zonder .build()).
 */
function createErrorCard(message) {
    var card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader().setTitle('Fout Opgetreden'));
    card.addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("Er is een onverwachte fout opgetreden:\n" + message)));
    // Optioneel: Knop om opnieuw te proberen de homepage te laden
    card.addSection(CardService.newCardSection().addWidget(
        CardService.newTextButton().setText("Opnieuw Laden")
        .setOnClickAction(CardService.newAction().setFunctionName("buildHomepageCard"))
    ));
    return card;
}

// --- Actie Functies (Worden aangeroepen door knoppen) ---

/**
 * Controleert of een Gmail label bestaat. Optioneel maakt het label aan als het niet bestaat.
 * @param {string} labelName De naam van het label om te controleren/maken.
 * @param {boolean} createIfNeeded Indien true, wordt het label aangemaakt als het niet bestaat.
 * @return {boolean} True als het label bestaat of succesvol is aangemaakt, anders false.
 */
function checkAndEnsureLabelExists(labelName, createIfNeeded) {
  if (!labelName || labelName.trim() === "") {
    Logger.log("checkAndEnsureLabelExists: Lege labelnaam ontvangen.");
    return false; // Lege naam is ongeldig
  }
  labelName = labelName.trim();

  try {
    var label = GmailApp.getUserLabelByName(labelName);
    if (label) {
      Logger.log("Label '" + labelName + "' bestaat al.");
      return true; // Label bestaat
    } else if (createIfNeeded) {
      // Label bestaat niet, maar we moeten het aanmaken
      Logger.log("Label '" + labelName + "' bestaat niet. Poging tot aanmaken...");
      try {
        label = GmailApp.createLabel(labelName);
        Logger.log("Label '" + labelName + "' succesvol aangemaakt.");
        return true; // Succesvol aangemaakt
      } catch (createErr) {
        // Fout bij aanmaken (bv. ongeldige tekens, te lang, conflict met systeemlabel)
        Logger.log("Fout bij aanmaken label '" + labelName + "': " + createErr);
        // We geven hier false terug, de aanroepende functie moet de fout afhandelen.
        return false;
      }
    } else {
      // Label bestaat niet en we mochten het niet aanmaken
      Logger.log("Label '" + labelName + "' bestaat niet (en createIfNeeded is false).");
      return false;
    }
  } catch (e) {
    // Algemene fout bij zoeken naar label
    Logger.log("Fout bij controleren/aanmaken label '" + labelName + "': " + e);
    return false;
  }
}

/**
 * Wordt aangeroepen door de "Reset Sheet Selectie" knop.
 * Verwijdert de opgeslagen sheet ID en vernieuwt de UI om opnieuw te zoeken/prompten.
 * @param {Object} e Event object.
 * @return {ActionResponse} Een response object om de UI te updaten.
 */
function handleClearSheetIdAction(e) {
  var userProperties = PropertiesService.getUserProperties();
  var currentSheetId = userProperties.getProperty(SHEET_ID_PROPERTY_KEY); // Lees huidige ID voor logging

  try {
    userProperties.deleteProperty(SHEET_ID_PROPERTY_KEY); // Verwijder de opgeslagen ID
    Logger.log("Gebruiker heeft Sheet ID reset uitgevoerd. Oude ID was: " + currentSheetId);

    // Roep buildHomepageCard opnieuw aan om de UI te vernieuwen.
    // Deze zal nu de findOrCreateOrPromptSheet logica uitvoeren.
    // buildHomepageCard geeft een array terug, we hebben de eerste kaart nodig.
    var refreshedHomepageCards = buildHomepageCard(e);
    var newCard = refreshedHomepageCards[0]; // Pak de eerste (en enige) kaart

    // Geef een ActionResponse terug om de huidige kaart te updaten
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(newCard)) // Update de UI met de nieuwe kaart
      .setNotification(CardService.newNotification().setText("Opgeslagen Sheet selectie gewist. Add-on zoekt opnieuw."))
      .build();

  } catch (err) {
    Logger.log("Fout tijdens uitvoeren handleClearSheetIdAction: " + err);
    // Geef een foutmelding terug als het verwijderen mislukt (onwaarschijnlijk)
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Fout bij resetten: " + err.message))
      .build();
  }
}

/**
 * Wordt aangeroepen door de 'Gebruik Deze Sheet' knop op de selectiekaart.
 * Valideert de gebruikersinvoer, slaat de ID op en toont de hoofdkaart.
 * @param {Object} e Event object met formulierinvoer.
 * @return {ActionResponse} Een response object voor de UI update.
 */
function saveSelectedSheet(e) {
  var userInput = e.formInput.sheet_input; // Haal de invoer van de gebruiker op
  var sheetId = null;
  var userProperties = PropertiesService.getUserProperties();

  if (!userInput) {
    // Geen invoer gegeven
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Voer een Sheet ID of URL in.'))
      .build();
  }

  // Probeer een geldige Sheet ID te extraheren uit de invoer (URL of directe ID)
  try {
    if (userInput.includes('/spreadsheets/d/')) {
      // Probeer ID uit URL te halen
      var match = userInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
          sheetId = match[1];
      }
    } else if (userInput.length > 20 && !userInput.includes(" ")) { // Ruwe check of het een ID zou kunnen zijn
      sheetId = userInput;
    }

    if (!sheetId) {
       throw new Error("Kon geen geldige Sheet ID vinden in de invoer.");
    }

    // Belangrijk: Valideer de ID door te proberen de sheet te openen
    SpreadsheetApp.openById(sheetId);
    Logger.log("Handmatig ingevoerde Sheet ID gevalideerd: " + sheetId);

    // Als validatie slaagt, sla de ID op in UserProperties
    userProperties.setProperty(SHEET_ID_PROPERTY_KEY, sheetId);

    // Bouw de hoofdactiekaart opnieuw, nu met de opgeslagen ID
    var mainCard = createMainActionCard(sheetId);
    // Gebruik updateCard om de huidige selectiekaart te vervangen door de hoofdkaart
    return CardService.newActionResponseBuilder()
         .setNavigation(CardService.newNavigation().updateCard(mainCard.build()))
         .setNotification(CardService.newNotification().setText('Sheet selectie succesvol opgeslagen!'))
         .build();

  } catch (err) {
    // Fout bij validatie (ongeldige ID, geen toegang, etc.)
    Logger.log("Fout bij opslaan/valideren handmatige sheet ID: " + err);
    // Blijf op de selectiekaart, maar toon een foutmelding aan de gebruiker
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Fout: Ongeldige ID/URL of geen toegang. Controleer de invoer en probeer opnieuw. (' + err.message + ')'))
      .build();
  }
}

/**
 * Wordt aangeroepen door de "Label Opslaan" of "Gebruik / Maak Dit Label" knop.
 * Controleert of het label bestaat, maakt het aan indien nodig, slaat het op
 * in UserProperties en vernieuwt naar de hoofdkaart.
 * @param {Object} e Event object met formulierinvoer.
 * @return {ActionResponse} Een response object om de UI te updaten.
 */
function handleSaveLabelAction(e) {
  var userProperties = PropertiesService.getUserProperties();
  var newLabelName = e.formInput.label_name_input;

  if (!newLabelName || newLabelName.trim() === "") {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Labelnaam mag niet leeg zijn."))
      .build();
  }
  newLabelName = newLabelName.trim();

  // Stap 1: Controleer of het label bestaat, maak het aan indien nodig
  var labelReady = checkAndEnsureLabelExists(newLabelName, true); // true = maak aan indien nodig

  if (!labelReady) {
    // Fout bij controleren of aanmaken (bv. ongeldige naam)
    Logger.log("handleSaveLabelAction: checkAndEnsureLabelExists mislukt voor '" + newLabelName + "'.");
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Kon label '" + newLabelName + "' niet gebruiken of aanmaken. Controleer de naam (geen speciale tekens zoals /)."))
      .build(); // Blijf op dezelfde kaart (setup of hoofdkaart instellingen)
  }

  // Stap 2: Label bestaat of is succesvol aangemaakt, sla nu op
  try {
    userProperties.setProperty(LABEL_NAME_PROPERTY_KEY, newLabelName);
    Logger.log("Labelnaam '" + newLabelName + "' succesvol opgeslagen in UserProperties.");

    // Stap 3: Haal sheet ID op en toon de hoofdkaart
    var sheetId = userProperties.getProperty(SHEET_ID_PROPERTY_KEY);
    if (!sheetId) {
      // Dit zou niet mogen gebeuren, maar vang het op
      Logger.log("Fout: Sheet ID niet gevonden na opslaan label.");
      // Probeer de homepage opnieuw op te bouwen, die zal de sheet setup opnieuw triggeren
       var errorCard = createErrorCard("Sheet configuratie verloren. Add-on wordt opnieuw geladen.").build();
       return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().updateCard(errorCard))
          .setNotification(CardService.newNotification().setText("Label opgeslagen, maar sheet moet opnieuw ingesteld worden."))
          .build();
    }

    // Bouw de hoofdkaart opnieuw met de (mogelijk nieuwe) label info
    var mainCard = createMainActionCard(sheetId).build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(mainCard)) // Ga naar de hoofdkaart
      .setNotification(CardService.newNotification().setText("Label '" + newLabelName + "' ingesteld."))
      .build();

  } catch (err) {
    Logger.log("Fout tijdens opslaan label property of herbouwen kaart: " + err);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Fout bij opslaan instelling: " + err.message))
      .build();
  }
}


/**
 * Wordt aangeroepen door de 'Verwerk Nu' knop.
 * Voert de daadwerkelijke logica uit om emails te scannen (met geconfigureerd label),
 * data te extraheren en naar de geselecteerde spreadsheet te schrijven.
 * @param {Object} e Event object.
 * @return {ActionResponse} Een response object met de resultaatkaart.
 */
function runProcessingLogic(e) {
  var userProperties = PropertiesService.getUserProperties();
  var sheetId = userProperties.getProperty(SHEET_ID_PROPERTY_KEY);
  var configuredLabelName = getConfiguredLabelName();
  var card;
  var lastProcessedDraftId = null; 

  // Extra veiligheidscheck: hebben we wel een ID om mee te werken?
  if (!sheetId) {
      Logger.log("Fout: runProcessingLogic aangeroepen zonder geldige sheet ID in properties.");
      card = createErrorCard("Kon de doel-spreadsheet niet bepalen. Probeer de add-on opnieuw te laden via de zijbalk.");
      return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().updateCard(card.build()))
          .build();
  }

  // --- Start van de verwerkingslogica ---
  var userEmail = Session.getActiveUser().getEmail();
  var rowsToAdd = []; // Array om alle rijen te verzamelen voor bulk-schrijven
  var processedCount = 0; // Teller voor verwerkte emails
  var errorMsg = null; // Variabele om eventuele fouten op te slaan
  var resultMessage = ""; // Boodschap voor de succeskaart

  // Helper functie om e-mailadressen te extraheren uit een string (bv. "Naam <a@b.com>, c@d.com")
  function extractEmails(emailString) {
      if (!emailString) return [];
      var emailRegex = /[\w\.\+\-]+@[\w\.\-]+\.\w+/gi;
      var emails = emailString.match(emailRegex);
      return emails ? [...new Set(emails)] : [];
  }

  // NIEUW: Helper functie om naam te extraheren uit "Naam <email@adres.com>"
  function extractName(nameAndEmailString) {
      if (!nameAndEmailString) return "";
      var name = nameAndEmailString; // Standaard de hele string
      var emailStartIndex = nameAndEmailString.indexOf('<');
      if (emailStartIndex > 0) { // Zorg ervoor dat '<' niet het eerste teken is
          name = nameAndEmailString.substring(0, emailStartIndex).trim();
      }
      // Verwijder eventuele dubbele aanhalingstekens rond de naam
      name = name.replace(/^"|"$/g, '');
      // Als de naam leeg is na trimmen/vervangen, gebruik dan de volledige string als fallback
      return name || nameAndEmailString;
  }


  try {
    // 1. Open de Google Sheet via de opgeslagen/gevalideerde ID
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(TARGET_TAB_NAME);
     // Zorg dat het doel-tabblad bestaat, maak aan indien nodig
     if (!sheet) {
       sheet = ss.insertSheet(TARGET_TAB_NAME);
       sheet.appendRow(["Datum", "Afzender", "Ontvanger", "Titel", "Beschrijving", "Email Link", "Attachment/Link"]);
       Logger.log("Tabblad '" + TARGET_TAB_NAME + "' was verdwenen en is opnieuw aangemaakt.");
     }

    // 2. Zoek het geconfigureerde Gmail label
    Logger.log("Zoeken naar Gmail label: '" + configuredLabelName + "'");
    var label = GmailApp.getUserLabelByName(configuredLabelName);
     if (!label) {
         throw new Error("Gmail label '" + configuredLabelName + "' niet gevonden. Controleer de naam in de instellingen en in Gmail.");
     }

    // 3. Haal alle threads op met dit label
    var threads = label.getThreads();
     Logger.log("Aantal threads gevonden met label '" + configuredLabelName + "': " + threads.length);

    // 4. Loop door elke thread en elk bericht binnen die thread
     for (var i = 0; i < threads.length; i++) { // --- Buitenste loop (per thread) ---
       var currentThread = threads[i];
       var messages = currentThread.getMessages(); // Haal alle berichten in deze thread op

       for (var j = 0; j < messages.length; j++) { // --- Binnenste loop (per bericht) ---
         var message = messages[j];
         processedCount++; // Tel elk bericht dat we bekijken

         // Haal basisinformatie op
         var msgDate = message.getDate();
         var msgSubject = message.getSubject(); // Onderwerp ophalen
         var msgSenderRaw = message.getFrom(); // Bevat vaak naam + email
         var senderName = extractName(msgSenderRaw); // Probeer naam te extraheren
         // Ontvanger: Combineer To en Cc, extraheer unieke ontvanger
         // LET OP: We gebruiken nog steeds e-mailadressen voor ontvangers omdat
         // het betrouwbaar extraheren van namen voor *elke* ontvanger uit To/Cc complex is.
         // De kolom "Ontvanger" bevat dus het e-mailadres van de specifieke ontvanger voor die rij.
         var recipientString = [message.getTo(), message.getCc()].filter(Boolean).join(',');
         var recipientEmails = extractEmails(recipientString); // Geeft array van unieke emails

         // Genereer de link naar de e-mail
         var threadId = currentThread.getId(); // threadId van de huidige thread
         var messageId = message.getId();      // messageId van het specifieke bericht
         var emailLink = "https://mail.google.com/mail/u/0/#all/" + threadId + "/" + messageId;
         

         // --- Verwerk Attachments ---
         var attachments = message.getAttachments();
         for (var k = 0; k < attachments.length; k++) {
           var attachmentName = attachments[k].getName();
           if (attachmentName) {
             // Loop door elke ontvanger voor dit attachment
             for (var r = 0; r < recipientEmails.length; r++) {
               var recipient = recipientEmails[r]; // Dit is het E-MAILADRES van de ontvanger
               // AANGEPAST: Voeg rij toe met nieuwe structuur
               rowsToAdd.push([
                   msgDate,          // Datum
                   senderName,       // Afzender (Naam/Raw)
                   recipient,        // Ontvanger (Email)
                   msgSubject,       // Titel
                   "",               // Beschrijving (Leeg)
                   emailLink,        // Email Link
                   attachmentName    // Attachment/Link
               ]);
             }
           }
         } // Einde loop attachments

         // --- Verwerk Links ---
         var body = message.getPlainBody(); // Haal platte tekst op voor link extractie
         var urlRegex = /(https?:\/\/[^\s"'<>`]+)/gi;
         var links = body.match(urlRegex);
         if (links) {
           var uniqueLinks = new Set(links); // Voorkom dubbele links uit dezelfde email body
           uniqueLinks.forEach(function(link) {
             // Loop door elke ontvanger voor deze link
             for (var r = 0; r < recipientEmails.length; r++) {
                var recipient = recipientEmails[r]; // Dit is het E-MAILADRES van de ontvanger
                // AANGEPAST: Voeg rij toe met nieuwe structuur
                rowsToAdd.push([
                    msgDate,          // Datum
                    senderName,       // Afzender (Naam/Raw)
                    recipient,        // Ontvanger (Email)
                    msgSubject,       // Titel
                    "",               // Beschrijving (Leeg)
                    emailLink,        // Email Link
                    link              // Attachment/Link
                ]);
             }
           });
         } // Einde verwerking links

        // Check of dit bericht een concept is en sla de message ID op
         if (message.isDraft()) {
             lastProcessedDraftId = message.getId();
             Logger.log("Bericht " + lastProcessedDraftId + " is een concept. ID opgeslagen.");
         } else {
             lastProcessedDraftId = null;
         } // einde concept detectie

       } // --- Einde binnenste loop (messages) ---

       // Verwijder het label van de HUIDIGE thread NADAT alle berichten zijn verwerkt.
       try {
           currentThread.removeLabel(label);
           Logger.log("Label '" + configuredLabelName + "' verwijderd van thread " + currentThread.getId());
       } catch (labelErr) {
           Logger.log("Fout bij verwijderen label '" + configuredLabelName + "' van thread " + currentThread.getId() + ": " + labelErr);
       }

    } // --- Einde buitenste loop (threads) ---

    // 5. Schrijf alle verzamelde data in één keer naar de Sheet
    if (rowsToAdd.length > 0) {
       var firstDataRow = sheet.getLastRow() + 1;
       // Correctie voor lege sheet of sheet met alleen header (check blijft op A1 = "Datum")
       if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && sheet.getRange("A1").getValue() === "Datum")) {
            firstDataRow = sheet.getLastRow() + 1;
       }
       // Zorg ervoor dat de range breedte overeenkomt met het aantal kolommen in rowsToAdd
       sheet.getRange(firstDataRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
       Logger.log(rowsToAdd.length + " rijen toegevoegd aan sheet '" + sheet.getParent().getName() + "', tabblad '" + sheet.getName() + "'.");
    }

    // Stel de succesboodschap samen
    resultMessage = "Verwerking voltooid voor label '<b>" + configuredLabelName + "</b>'.\n" + 
                    "E-mails onderzocht: " + processedCount + "\n" +
                    "Rijen (items * ontvangers) toegevoegd: " + rowsToAdd.length;

  } catch (err) {
    // Vang eventuele fouten tijdens de verwerking op
    Logger.log("Error tijdens verwerking: " + err + "\nStack: " + err.stack);
    errorMsg = "Fout opgetreden tijdens verwerking: " + err.message;
  }

  // 6. Bouw de resultaatkaart (succes of fout)
  card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Verwerkingsresultaat'));
  var responseSection = CardService.newCardSection();
  if (errorMsg) {
      responseSection.addWidget(CardService.newTextParagraph().setText("<b>Fout:</b>\n" + errorMsg));
  } else {
      responseSection.addWidget(CardService.newTextParagraph().setText(resultMessage));
      // --- VOEG VERZENDKNOP TOE INDIEN VAN TOEPASSING ---
      if (lastProcessedDraftId) {
          Logger.log("Concept ID " + lastProcessedDraftId + " gevonden, knop 'Verzend Concept Nu' wordt toegevoegd.");
          responseSection.addWidget(CardService.newButtonSet()
              .addButton(CardService.newTextButton()
                  .setText("Verzend Concept Nu")
                  .setOnClickAction(CardService.newAction()
                      .setFunctionName("sendDraftAction")
                      .setParameters({'draftMessageId': lastProcessedDraftId}))
                  .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              )
          );
      }
      // --- EINDE VERZENDKNOP CODE ---
  }
  // Knop om terug te gaan naar de hoofdkaart
  responseSection.addWidget(CardService.newTextButton().setText("Terug")
      .setOnClickAction(CardService.newAction().setFunctionName("buildHomepageCard")));
  card.addSection(responseSection);

  // Update de UI met de resultaatkaart
  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(card.build()))
      .build();
} // --- Einde runProcessingLogic functie ---

/**
 * Wordt aangeroepen door de "Verzend Concept Nu" knop.
 * Zoekt het juiste concept op basis van de message ID en probeert het te verzenden.
 * @param {Object} e Event object met parameters (incl. draftMessageId, wat eigenlijk de message ID is).
 * @return {ActionResponse} Response met alleen een notificatie (succes of fout).
 */
function sendDraftAction(e) {
  // De ID die we meekrijgen is de ID van het GmailMessage object van het concept
  var messageIdOfDraft = e.parameters.draftMessageId;
  Logger.log("sendDraftAction aangeroepen voor message ID: " + messageIdOfDraft);

  if (!messageIdOfDraft) {
    Logger.log("Fout: draftMessageId (message ID) parameter ontbreekt in sendDraftAction.");
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Fout: Kon concept ID niet vinden."))
      .build();
  }

  var draftFound = false; // Flag om bij te houden of we de draft vinden

  try {
    // Haal ALLE concepten op
    var allDrafts = GmailApp.getDrafts();
    Logger.log("Aantal concepten gevonden in totaal: " + allDrafts.length);

    // Loop door alle concepten om de juiste te vinden
    for (var i = 0; i < allDrafts.length; i++) {
      var draft = allDrafts[i];
      var message = null;
      var currentMessageId = null;

      try {
          // Haal het bericht op dat bij dit concept hoort
          message = draft.getMessage();
          if (message) {
              currentMessageId = message.getId();
          }
      } catch (getMsgErr) {
          // Kan gebeuren bij zeer oude of corrupte drafts
          Logger.log("Kon bericht niet ophalen voor draft index " + i + ": " + getMsgErr);
          continue; // Ga naar de volgende draft
      }


      // Vergelijk de ID van het bericht van de draft met de ID die we zoeken
      if (message && currentMessageId === messageIdOfDraft) {
        Logger.log("Overeenkomende draft gevonden! Draft ID: " + draft.getId() + ", Message ID: " + currentMessageId);
        draftFound = true; // We hebben hem gevonden!

        // Probeer te verzenden
        draft.send();
        Logger.log("Concept met message ID " + messageIdOfDraft + " succesvol verzonden.");

        // Geef succesmelding terug en stop de loop/functie
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText("Concept succesvol verzonden!"))
          .build();
      }
    } // Einde loop door allDrafts

    // Als we hier komen, is de loop voltooid ZONDER de draft te vinden
    if (!draftFound) {
      Logger.log("Fout: Kon geen overeenkomende draft vinden voor message ID " + messageIdOfDraft + " (mogelijk al verzonden/verwijderd).");
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Fout bij verzenden: Concept niet gevonden (mogelijk al verzonden of verwijderd)."))
        .build();
    }

  } catch (err) {
    // Vang algemene fouten tijdens het proces op
    Logger.log("Algemene fout bij verzenden concept met message ID " + messageIdOfDraft + ": " + err + "\nStack: " + err.stack);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Fout bij verzenden: " + err.message))
      .build();
  }
}
