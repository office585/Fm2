function saveAndTrashMewsAndSzamlazzReports() {
  // --- BEÁLLÍTÁSOK ---
  var mewsFolderId = "1KGd5i9yH9UxJw6yTSveZBpbwwaS3Hj6_"; 
  var szamlazzFolderId = "16KYmZhM-F08ZNHj-3VQf1TZXszHki1Cf";
  var fmatracFolderId = "19bo5GiU6lrPEgfrSbJrO74e7pWD9Ng7U"; // Az új FMATRAC mappa
  
  // Szűrőfeltételek listája
  var filters = [
    {
      subject: 'subject:"Scheduled export of Payment report"',
      folderId: mewsFolderId,
      name: "Mews Report"
    },
    {
      subject: 'subject:"Számlázz.hu ÁFA export"', 
      folderId: szamlazzFolderId,
      name: "Számlázz.hu ÁFA"
    },
    {
      // Megkeresi az összes e-mailt, ami ezzel kezdődik
      subject: 'subject:"FMATRAC Teljes Éves Riport"', 
      folderId: fmatracFolderId,
      name: "FMATRAC Riport"
    }
  ];
  // -------------------

  filters.forEach(function(config) {
    var folder;
    try {
      folder = DriveApp.getFolderById(config.folderId);
    } catch (e) {
      Logger.log("Hiba: A mappa nem található (" + config.name + "): " + config.folderId);
      return; // Ugrik a következő szűrőre
    }

    // Keresés: csak a beérkező levelek között (label:inbox)
    var searchTerm = config.subject + ' label:inbox';
    var threads = GmailApp.search(searchTerm);

    Logger.log(config.name + " - Talált feldolgozandó szálak száma: " + threads.length);

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var attachments = message.getAttachments();
        
        for (var k = 0; k < attachments.length; k++) {
          var file = attachments[k];
          folder.createFile(file);
          Logger.log(config.name + " mentve: " + file.getName());
        }
      }
      
      // Miután minden csatolmányt kimentettünk az összes üzenetből a szálban, mehet a kukába
      threads[i].moveToTrash();
      Logger.log(config.name + " szál áthelyezve a kukába.");
    }
  });
}
