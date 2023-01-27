
/*************************************************************************************************************************
 *
 *                              Yes We Camp Planning Prod script
 * 
 * Ajout de fonctions au planning de production :
 *  - ajouter automatiquement les événements validés aux calendriers Google, selon les 
 *    conditions précisées dans l'onglet LIEUX du classeur.
 *  - création du rapport hebdomadaire des événements dans un document texte sur le drive
 *  - archivage des événements passés
 * 
 * version : 2.0
 * date : 05/07/2021
 * authors : HJE, ARA, MRE
 * 
 * changelog :
 * - v1.9 - 21/03/2017 : - retrait des majuscules au mois et de l'année dans les dates textuelles
 *                       - ajout de tous les événements dans la génération de la partie "entre Buropolis" du rapport hebdo 
 * - v2.0 - 22/03/2017 : - gestion de la confirmation "Diffusé" permettant l'inscription aux calendriers publics & Buropolis
 *                       - Ajout de la protection des lignes ajoutées aux agendas 
 * - v2.1 - 28/03/2017 : - Corrections de l'export du rapport hebdo (événements diffusés, dates)
 * - v2.2 - 10/05/2018 : - Ajout gestion des couleurs des évènements
 * 
 *************************************************************************************************************************/


/* liste des agendas supplémentaires dans lesquels écrire
 * le reste est à voir au niveau de l'onglet LIEUX du document
 */
var CAL_NAME_BUROPOLIS = 'Buropolis - Programmation' // le nom de l'agenda dans lequel on ajoutera les events d'autres salles et les évènements réservés internes
var CAL_NAME_PUBLIC = 'Buropolis - Programmation' // le nom de l'agenda dans lequel seront publiés les events publics
var CAL_NAME_PILOTAGE = 'Buropolis - Programmation' // le nom de l'agenda dans lequel seront publiés les events lié au pilotage
// var CAL_NAME_BUROPOLIS = 'H - Pilotage' // le nom de l'agenda dans lequel on ajoutera les events d'autres salles et les évènements réservés internes
// var CAL_NAME_PUBLIC = 'H - Pilotage' // le nom de l'agenda dans lequel seront publiés les events publics
// var CAL_NAME_PILOTAGE = 'H - Pilotage' // le nom de l'agenda dans lequel seront publiés les events lié au pilotage



// constantes liées aux colonnes
var EVENT_IMPORTED = "Ajouté"; // Ajoutera le texte "AJOUTE" dans la colonne correspondante
var EVENT_RELEASED = "Diffusé"; // Ajoutera le texte "Diffusé" dans la colonne correspondante
var COL_CONFIRMATION = 0;
var COL_DATE = 1;
var COL_HORAIRE_DEB = 2;
var COL_HORAIRE_FIN = 3;
var COL_LIEU = 4;
var COL_TITRE = 5;
var COL_TYPE = 6;
var COL_PUBLIC = 7;
var COL_DETAIL = 8;
var COL_VOISIN = 9;
var COL_JAUGE = 10;
var COL_CONTACT_REF = 11;
var COL_CONTACT_EXT = 12;
var COL_CONTACT_TEL = 13;
var COL_AJOUTE = 14;

var ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
   var menuEntries = [{name: "Ajouter les événements à l'agenda", functionName: "exportToCalendar"}, {name: "Générer le rapport hebdomadaire", functionName: "generateWeeklyReport"}, {name: "Archiver les événements passés", functionName: "archiveData"}];
   ss.addMenu("YWC", menuEntries); // ajout du menu YWC dans le classeur
}

/******************************************************************************************************************************
 *              Exportation des évènements du classeur vers les calendriers
 ******************************************************************************************************************************/
function exportToCalendar() {        

    var currentSheet = SpreadsheetApp.getActiveSheet();
    var currentSheet = ss.getSheetByName('Planning');
    var firstRow = currentSheet.getActiveCell().getRow(); // ligne pointée par le curseur
    if (firstRow < 3 ) { firstRow = 3; } // sécurité pour pas cliquer avant la ligne 3
    var lastRow = currentSheet.getLastRow();   
    
 
    for (var i = firstRow ; i <= lastRow ; i++){
        // récupérer la ligne complète dans une variable / tableau, et faire les calculs à partir d'elle pour les performances
        var currentRange = currentSheet.getRange(i, 1, 1, 14).getValues(); 
        var currentLine = currentRange[0]; // le Range est un tableau à 2 dimensions, on sélectionne uniquement la première ligne                
        
        if (currentLine[COL_CONFIRMATION] == "Confirmé" && currentLine[COL_AJOUTE] != EVENT_IMPORTED){
            // on récupère la couleur de l'évènement dans l'onglet TYPES
            var eventColor = getColor(currentLine[COL_TYPE]);
              
            if (currentLine[COL_AJOUTE] == EVENT_RELEASED){ // confirmé & diffusé -> retirer diffusion publique                                
                var currentEvent = getEvent(currentLine);
                currentEvent.titre = currentLine[COL_TITRE]; 

                if (currentLine[COL_PUBLIC] == "PUBLIC EXT") {
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_PUBLIC);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }
                if (currentLine[COL_PUBLIC] == "buropolis") {
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_BUROPOLIS);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }
                if (currentLine[COL_PUBLIC] == "OCCUPANTS"){
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_BUROPOLIS);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }
                //unprotectLine(i, currentSheet);
            } else { // confirmé && rien -> ajout calendriers non-publics
                var currentEvent = getEvent(currentLine);
                var calTab = getCalendars(currentEvent.location); // récupère le nom des 3 potentiels calendriers sur lesquels diffuser
                if (calTab != null){
                    if (calTab[0][0] != ''){ // diffusion agenda1
                        addEventToCalendar(currentEvent, calTab[0][0], false, eventColor);                        
                    }
                    if (calTab[0][1] != ''){ // diffusion agenda2
                        addEventToCalendar(currentEvent, calTab[0][1], false, eventColor);
                    }
                    if (calTab[0][2] != ''){ // diffusion publique, info moindre
                        addEventToCalendar(currentEvent, calTab[0][2], true, eventColor);
                    }
                    currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on marque l'event comme ajouté
                    SpreadsheetApp.flush();
                }
                if (currentLine[COL_PUBLIC] == "pilotage") { // event de pilotage, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_PILOTAGE, false, eventColor);                    
                }                
                
              //protectLine(i, currentSheet);
            }
        } else if (currentLine[COL_CONFIRMATION] == "Diffusé" && currentLine[COL_AJOUTE] != EVENT_RELEASED){
            // on récupère la couleur de l'évènement dans l'onglet TYPES
            var eventColor = getColor(currentLine[COL_TYPE]);
            
            var currentEvent = getEvent(currentLine);
            if (currentLine[COL_AJOUTE] == EVENT_IMPORTED){ // diffusé & ajouté -> ajout calendriers publics
                if (currentLine[COL_PUBLIC] == "PUBLIC EXT") { // si l'event est public on l'ajoute à un calendrier public avec une description moindre             
                    addEventToCalendar(currentEvent, CAL_NAME_PUBLIC, true, eventColor);                    
                }
                if (currentLine[COL_PUBLIC] == "buropolis") { // si l'event est spécifique aux buropolis, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_BUROPOLIS, false, eventColor);                    
                }
                if (currentLine[COL_PUBLIC] == "OCCUPANTS") { // si l'event est spécifique aux buropolis, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_BUROPOLIS, false, eventColor);                    
                }
                currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_RELEASED); // on marque l'event comme diffusé
                SpreadsheetApp.flush();
                //protectLine(i, currentSheet);
            } else { // diffusé & rien -> ajout calendriers publics & non-publics
                var calTab = getCalendars(currentEvent.location); // récupère le nom des 3 potentiels calendriers sur lesquels diffuser
                if (calTab != null){
                    if (calTab[0][0] != ''){ // diffusion agenda1
                        addEventToCalendar(currentEvent, calTab[0][0], false, eventColor);                        
                    }
                    if (calTab[0][1] != ''){ // diffusion agenda2
                        addEventToCalendar(currentEvent, calTab[0][1], false, eventColor);
                    }
                    if (calTab[0][2] != ''){ // diffusion publique, info moindre
                        addEventToCalendar(currentEvent, calTab[0][2], true, eventColor);
                    }                    
                }
                if (currentLine[COL_PUBLIC] == "pilotage") { // event de pilotage, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_PILOTAGE, false, eventColor);                    
                }
                if (currentLine[COL_PUBLIC] == "PUBLIC EXT") { // si l'event est public on l'ajoute à un calendrier public avec une description moindre             
                    addEventToCalendar(currentEvent, CAL_NAME_PUBLIC, true, eventColor);                    
                }
                if (currentLine[COL_PUBLIC] == "buropolis") { // si l'event est spécifique aux buropolis, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_BUROPOLIS, false, eventColor);                    
                }
                if (currentLine[COL_PUBLIC] == "OCCUPANTS") { // si l'event est spécifique aux buropolis, on l'ajoute au calendrier correspondant 
                    addEventToCalendar(currentEvent, CAL_NAME_BUROPOLIS, false, eventColor);                    
                }
                currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_RELEASED); // on marque l'event comme diffusé
                SpreadsheetApp.flush();
                //protectLine(i, currentSheet);
            }   
        } else if (currentLine[COL_CONFIRMATION] == "Reporté" || currentLine[COL_CONFIRMATION] == "Annulé"){
            if (currentLine[COL_AJOUTE] == EVENT_RELEASED){ // annulé & diffusé -> retirer calendriers publics & non-publics                
                var currentEvent = getEvent(currentLine);
                currentEvent.titre = currentLine[COL_TITRE];

                if (currentLine[COL_PUBLIC] == "PUBLIC EXT") {
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_PUBLIC);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }
                if (currentLine[COL_PUBLIC] == "buropolis") {
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_BUROPOLIS);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }
                if (currentLine[COL_PUBLIC] == "OCCUPANTS"){
                    var deleted = deleteEventFromCalendar(currentEvent, CAL_NAME_BUROPOLIS);
                    if (deleted){
                        currentSheet.getRange(i, COL_AJOUTE+1).setValue(EVENT_IMPORTED); // on enlève l'event comme diffusé
                    }
                }

                var calTab = getCalendars(currentLine[COL_LIEU]);

                if (calTab != null){
                    if (calTab[0][0] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][0]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                    if (calTab[0][1] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][1]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                    if (calTab[0][2] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][2]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                }
                if (currentLine[COL_PUBLIC] == "pilotage") {
                    deleteEventFromCalendar(currentEvent, CAL_NAME_PILOTAGE);                    
                }    
                //unprotectLine(i, currentSheet);
            } else if (currentLine[COL_AJOUTE] == EVENT_IMPORTED){ // annulé & ajouté -> retirer calendriers non-publics                
                var currentEvent = getEvent(currentLine);
                currentEvent.titre = currentLine[COL_TITRE];

                var calTab = getCalendars(currentLine[COL_LIEU]);

                if (calTab != null){
                    if (calTab[0][0] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][0]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                    if (calTab[0][1] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][1]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                    if (calTab[0][2] != ''){
                        var deleted = deleteEventFromCalendar(currentEvent, calTab[0][2]);
                        if (deleted){
                            currentSheet.getRange(i, COL_AJOUTE+1).setValue(""); // on enlève l'event comme ajouté
                        }
                    }
                }
                if (currentLine[COL_PUBLIC] == "pilotage") {
                    deleteEventFromCalendar(currentEvent, CAL_NAME_PILOTAGE);                    
                }      
                //unprotectLine(i, currentSheet);             
            }
        }                              
    }
}

/*
 * renvoie le(s) calendrier(s) dans le(s)quel(s) on va écrire selon la localisation de l'event 
 * va chercher cette information dans l'onglet LIEUX du classeur 
 * renvoie un tableau de taille 3, positionnel, avec les noms des calendriers 
 */
function getCalendars(location) {
    var locationSheet = ss.getSheetByName('LIEUX');
    
    var tempTab = locationSheet.getRange(3, 1, locationSheet.getLastRow()-3).getValues();
    
    var locationLine = 0;
    for (var i = 0 ; i < tempTab.length ; i++){
        if (location == tempTab[i]){
            locationLine = i+3;
            break;
        }
    }

    if (locationLine != 0) {
        return locationSheet.getRange(locationLine, 3, 1, 3).getValues();
    } else {
        return null;
    }
}

/*
 * renvoie la couleur associée au type de l'évènement
 * va chercher cette info dans l'onglet TYPES du classeur
 */
function getColor(typeEvent) {
  var colorSheet = ss.getSheetByName('TYPES');
  var tempTab = colorSheet.getRange(2, 1, colorSheet.getLastRow()-2).getValues();
  var typeLine = 0;
  for (var i = 0 ; i < tempTab.length ; i++){
    if (typeEvent == tempTab[i]) {
      typeLine = i+2;
      break;
    }
  }
 
 //if (currentColor.match(/^(PALE_BLUE|PALE_GREEN|MAUVE|PALE_RED|YELLOW|ORANGE|CYAN|GRAY|BLUE|GREEN|RED)$/))
 if (typeLine != 0){
    try {
      var currentColor = eval('COLORS.'+colorSheet.getRange(typeLine, 2, 1, 1).getValue());
      return currentColor;
    } catch (e) {
      return null;
    }
  }
  return null;
}

/*
 *  Renvoie un objet événement formaté depuis une ligne du classeur sous forme de tableau
 */
function getEvent(tabRange){
     // récupération et formatage des dates et heures depuis la feuille de calcul
    var cellDate = tabRange[COL_DATE];
    var cellHeureDeb = tabRange[COL_HORAIRE_DEB];
    var cellHeureFin = tabRange[COL_HORAIRE_FIN];
    var jour = cellDate.getDate();
    var mois = cellDate.getMonth();
    var annee = cellDate.getFullYear();
    var heuresDebut = getNbHours(cellHeureDeb);
    var minutesDebut = getNbMinutes(cellHeureDeb);
    
    // pour éviter des problèmes de création si l'heure de fin est vide
    if (tabRange[COL_HORAIRE_FIN] != ""){
        var heuresFin = getNbHours(cellHeureFin);
        var minutesFin = getNbMinutes(cellHeureFin);
    } else {
        var heuresFin = heuresDebut + 1;
        var minutesFin = minutesFin;
    }
    
    var dateDebut = new Date(annee, mois, jour, heuresDebut, minutesDebut, 0, 0);
    
    // si l'heure de fin est < à l'heure de début, création à j+1
    if (heuresFin > heuresDebut) {
        var dateFin = new Date(annee, mois, jour, heuresFin, minutesFin, 0, 0);
    } else {
        var dateFin = new Date(annee, mois, jour+1, heuresFin, minutesFin, 0, 0);
    }
    
    
    
    // création du titre de l'event
    var type = tabRange[COL_TYPE];
    var titreStr = "";
    if (type != "") {
        titreStr = " " + type + "  ";
    } 
    
    var titre = tabRange[COL_TITRE];
    if (titre == "") { // si la colonne F est vide, le titre est la colonne I
        titre = tabRange[COL_DETAIL];
    }
    titreStr += titre;
    
    // création de la description de l'event
    var description = " || Détails : " + tabRange[COL_DETAIL]
                    + " || Structure porteuse : " + tabRange[COL_CONTACT_REF] + " - " + tabRange[COL_CONTACT_EXT] 
                    + " || Tel : " + tabRange[COL_CONTACT_TEL];

    var description_publique = " || Détails : " + tabRange[COL_DETAIL];

     // création de la localisation de l'event
    var location = tabRange[COL_LIEU];

    return {titre: titreStr, dateDebut: dateDebut, dateFin: dateFin, desc: description, publicDesc: description_publique, location: location};
}

/*
 *  ajoute un évènement au calendrier précisé, avec une description adaptée selon que l'évènement soit public ou non
 */
function addEventToCalendar(eventObj, calendarName, isPublic, color){
    var calendar = CalendarApp.getOwnedCalendarsByName(calendarName);
    
    if (calendar == "") {
      try {
        calendar = CalendarApp.getCalendarsByName(calendarName);
      } catch(e) {
        Logger.log("Calendrier " + calendarName + " non trouvé parmis les calendriers souscrits / créés");
      }
    }
    if (!isPublic){
        var calEvent = calendar[0].createEvent(eventObj.titre, eventObj.dateDebut, eventObj.dateFin, {description: eventObj.desc, location: eventObj.location});   
    } else {
        var calEvent = calendar[0].createEvent(eventObj.titre, eventObj.dateDebut, eventObj.dateFin, {description: eventObj.publicDesc, location: eventObj.location});   
    }
    if (color != null){
      calEvent.setColor(color);
    }
}

/*
 *  Modifie la couleur d'un évènement dans un calendrier précisé
 */
function changeEventColorFromCalendar(eventObj, calendarName, color){
    try{
        var calendar = CalendarApp.getOwnedCalendarsByName(calendarName);
        if (calendar == "") {
          try {
            calendar = CalendarApp.getCalendarsByName(calendarName);
          } catch(e) {
            Logger.log("Calendrier " + calendarName + " non trouvé parmis les calendriers souscrits / créés");
          }
        }
        var events = calendar[0].getEvents(eventObj.dateDebut, eventObj.dateFin, {search: eventObj.titre});
        if (events[0] != undefined){
          if (color != null){
            events[0].setColor(color);
            return true;
          }            
        }
        return false;
    } catch(e){
        Logger.log('problème de modification de la couleur de l\'event : ' + e.message + ' | fichier : ' + e.fileName + ' | ligne : ' + e.lineNumber);
    }
}

/*
 *  supprime un évènement du calendrier précisé
 */
function deleteEventFromCalendar(eventObj, calendarName){
    try{
        var calendar = CalendarApp.getOwnedCalendarsByName(calendarName);
        if (calendar == "") {
          try {
            calendar = CalendarApp.getCalendarsByName(calendarName);
          } catch(e) {
            Logger.log("Calendrier " + calendarName + " non trouvé parmis les calendriers souscrits / créés");
          }
        }
        var events = calendar[0].getEvents(eventObj.dateDebut, eventObj.dateFin, {search: eventObj.titre});
        if (events[0] != undefined){
            events[0].deleteEvent();
            return true;                        
        }
        return false;
    } catch(e){
        Logger.log('problème de suppression event : ' + e.message + ' | fichier : ' + e.fileName + ' | ligne : ' + e.lineNumber);
    }
}


 //  Protège la ligne i des modifications
function protectLine(i, sheet){    
    // protection de la ligne ajoutée aux calendriers (à l'exception de la colonne de validation)
    var rangeToProtect = sheet.getRange(i, 2, 1, 13);
    var name = (rangeToProtect.getValues()[0][4].substring(0, 5) + getTextDate(rangeToProtect.getValues()[0][0], true)).replace(/\s/g,"");         
    ss.setNamedRange(name, rangeToProtect); // on associe le range à un nom (sa description)

    var protection = rangeToProtect.protect().setDescription('Évènement validé').setRangeName(name);
    
    // ajout de la personne diffusant comme éditeur, 
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.addEditor('axel@yeswecamp.org');
    protection.addEditor('raphael@yeswecamp.org');
    protection.addEditor('leila@yeswecamp.org');
    protection.addEditor('julien.n@yeswecamp.org');
    protection.addEditor('rebecca.p@yeswecamp.org');
    protection.addEditor('fanelie@yeswecamp.org');
    //protection.removeEditors(protection.getEditors()); 
}
 

// enlève la protection de la ligne i

function unprotectLine(i, sheet){
    var rangeToUnprotect = sheet.getRange(i, 2, 1, 13);    
    
    var name = (rangeToUnprotect.getValues()[0][4].substring(0, 5) + getTextDate(rangeToUnprotect.getValues()[0][0], true)).replace(/\s/g,"");
 
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    for (var i = 0; i < protections.length; i++) {
        var protection = protections[i];
        if (protection.getRangeName() == name){
            protection.remove();
        }
    }
}

 
/******************************************************************************************************************************
 *              Archivage des évènements passés
 ******************************************************************************************************************************/
function archiveData() { 
  // ouverture du document d'archivage 
    try {    
        var filesIt = DriveApp.getFilesByName("Buropolis - Planning programmation");        
        if (filesIt.hasNext()){
            var Archives = SpreadsheetApp.openById(filesIt.next().getId());                 
            var archiveSheet = Archives.getSheetByName("Archives");                   
        }  
    } catch (ex) {        
        Logger.log('Problème d\'ouverture du document d\'archivage ' + ex.fileName + ' raison : ' + ex.message);
    }

    var currentSheet = SpreadsheetApp.getActiveSheet();
    var firstRow = currentSheet.getActiveCell().getRow(); // ligne pointée par le curseur
    if (firstRow < 3 ) { firstRow = 3; } // sécurité pour pas cliquer avant la ligne 5
    var lastRow = currentSheet.getLastRow();     
   
    for (var i = firstRow ; i <= lastRow ; i++){
        // récupérer la ligne complète dans une variable / tableau, et faire les calculs à partir d'elle pour les performances
        var currentRange = currentSheet.getRange(i, 1, 1, 14).getValues(); 

        // dans le cas où la date est passée, on archive dans un document externe
        var dateDuJour = new Date();
        var dateLigne = currentRange[0][COL_DATE];
        
        if  ((dateDuJour.getUTCFullYear() == dateLigne.getUTCFullYear() && dateDuJour.getUTCMonth() == dateLigne.getUTCMonth() && dateDuJour.getUTCDate() > dateLigne.getUTCDate()+1)
            || (dateDuJour.getUTCFullYear() == dateLigne.getUTCFullYear() && dateDuJour.getUTCMonth() > dateLigne.getUTCMonth())
            || (dateDuJour.getUTCFullYear() > dateLigne.getUTCFullYear())){
                        
            // attention, les heures sont modifiées (-10minutes...) il faut les remodifier...
            currentRange[0][COL_HORAIRE_DEB] = new Date(currentRange[0][COL_DATE].getUTCFullYear(), currentRange[0][COL_DATE].getUTCMonth(), currentRange[0][COL_DATE].getUTCDate(), getNbHours(currentRange[0][COL_HORAIRE_DEB]), getNbMinutes(currentRange[0][COL_HORAIRE_DEB]), 0, 0);
            currentRange[0][COL_HORAIRE_FIN] = new Date(currentRange[0][COL_DATE].getUTCFullYear(), currentRange[0][COL_DATE].getUTCMonth(), currentRange[0][COL_DATE].getUTCDate(), getNbHours(currentRange[0][COL_HORAIRE_FIN]), getNbMinutes(currentRange[0][COL_HORAIRE_FIN]), 0, 0);
            
            archiveSheet.getRange(archiveSheet.getLastRow()+1, 1, 1, 14).setValues(currentRange);
            // Suppression de la ligne courante, attention !
            currentSheet.deleteRow(i);
            i--;
            lastRow--;            
            SpreadsheetApp.flush();        
        }

        if (dateLigne.getTime() > dateDuJour.getTime()){ // si la date dépasse celle du jour, on arrête le script
            break;
        }
    }  
}
/******************************************************************************************************************************
 *              Génération d'un document récapitulatif des évènements de la semaine
 ******************************************************************************************************************************/
function generateWeeklyReport(){
    var currentSheet = SpreadsheetApp.getActiveSheet();
    var firstRow = currentSheet.getActiveCell().getRow(); // ligne pointée par le curseur 
    var lastRow = currentSheet.getLastRow();
    var dateRapport = currentSheet.getRange(firstRow, COL_DATE+1).getValue();
    var dateFin = new Date(dateRapport.getUTCFullYear(), dateRapport.getUTCMonth(), dateRapport.getUTCDate() + 7); // date 7j plus tard pour fin du rapport        
    var docName = "rapport " + dateRapport.toDateString();
    
    // ouverture du document
    try {
        // on vérifie si le fichier existe déjà et on l'ouvre sinon on le crée
        var filesCollection = DriveApp.getFilesByName(docName);
        if (filesCollection.hasNext()){
            var doc = DocumentApp.openById(filesCollection.next().getId());                        
            //Logger.log('fichier trouvé, on l\'ouvre');
            doc.getBody().clear(); // on clear le contenu afin de le recréer                                   
        } else {
            var doc = DocumentApp.create(docName);
        }                 
    } catch (ex) {
        Logger.log('Problème d\'ouverture ou de création du document ' + ex.fileName + ' raison : ' + ex.message);
    }

    // le contenu du document
    var body = doc.getBody();     
 
    // on écrit une fois la date, et on sauvegarde sa valeur
    var currentDate = currentSheet.getRange(firstRow, COL_DATE+1).getValue();
    var saveMonth = currentDate.getUTCMonth();
    var saveDay = currentDate.getUTCDate();
    
    var textDate = getTextDate(currentDate, false); 
    paraDate = body.appendParagraph(textDate); 
    paraDate.setAttributes(STYLE_TITRE1);

    // tableaux servant à sauvegarder les textes des expos, des cours et des buropolis, que l'on ajoutera en fin de document
    var tabExpos = ["Expositions en cours"];        
    var tabCoursEtStages = ["Les cours et stages"];
    var tabGV = ["Entre Buropolis"];

    for (var i = firstRow ; i <= lastRow ; i++){        
        // on sélectionne les évènements confirmés
        if ((currentSheet.getRange(i, COL_CONFIRMATION+1).getValue() == "Confirmé" || currentSheet.getRange(i, COL_CONFIRMATION+1).getValue() == "Diffusé") && currentSheet.getRange(i, COL_PUBLIC+1).getValue() == "PUBLIC EXT"){  // événements tout public          
            // si le type d'évènement est EXPO, on l'ajoute au paragraphe expos, que l'on appendera en toute fin, pareil pour COURS ET STAGES, sinon on ajoute au fur et à mesure            
            if (currentSheet.getRange(i, COL_TYPE+1).getValue() == "EXPO"){                                
                // écrire le titre de l'expo
                var objExpo = {};                
                objExpo.titre = currentSheet.getRange(i, COL_TITRE+1).getValue();
                
                // écrire la date & heure & lieu
                var dateEvent = getTextDate(currentSheet.getRange(i, COL_DATE+1).getValue(), false);
                // var heureDeb = currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10));
                // var heureFin = currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10));
                var tempDateDeb = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10, 0, 0);                
                var tempDateFin = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10, 0, 0);
                var heureDeb = tempDateDeb.getHours() + "h" + (tempDateDeb.getUTCMinutes()<10?"0"+tempDateDeb.getUTCMinutes() : tempDateDeb.getUTCMinutes());
                var heureFin = tempDateFin.getHours() + "h" + (tempDateFin.getUTCMinutes()<10?"0"+tempDateFin.getUTCMinutes() : tempDateFin.getUTCMinutes());
                objExpo.desc = dateEvent + " - " + heureDeb + "-" + heureFin + " - " + currentSheet.getRange(i, COL_LIEU+1).getValue();     

                // écrire le détail
                objExpo.info = currentSheet.getRange(i, COL_DETAIL+1).getValue();                 

                // ajout au tableau des expos
                tabExpos.push(objExpo);

            } else if (currentSheet.getRange(i, COL_TYPE+1).getValue() == "COURS ET STAGES"){                
                var objCours = {};
                // écrire le titre du cours
                objCours.titre = currentSheet.getRange(i, COL_TITRE+1).getValue();
                
                // écrire la date & heure & lieu
                var dateEvent = getTextDate(currentSheet.getRange(i, COL_DATE+1).getValue(), false);
                // var heureDeb = currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10));
                // var heureFin = currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10));
                var tempDateDeb = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10, 0, 0);                
                var tempDateFin = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10, 0, 0);
                var heureDeb = tempDateDeb.getHours() + "h" + (tempDateDeb.getUTCMinutes()<10?"0"+tempDateDeb.getUTCMinutes() : tempDateDeb.getUTCMinutes());
                var heureFin = tempDateFin.getHours() + "h" + (tempDateFin.getUTCMinutes()<10?"0"+tempDateFin.getUTCMinutes() : tempDateFin.getUTCMinutes());
                objCours.desc = dateEvent + " - " + heureDeb + "-" + heureFin + " - " + currentSheet.getRange(i, COL_LIEU+1).getValue();
               
                // écrire le détail
                objCours.info = currentSheet.getRange(i, COL_DETAIL+1).getValue();               

                // ajout au tableau des cours                
                tabCoursEtStages.push(objCours);

            } else {     
                // écrire la date en toutes lettres si elle a changé
                var currentMonth = currentSheet.getRange(i, COL_DATE+1).getValue().getUTCMonth();
                var currentDay = currentSheet.getRange(i, COL_DATE+1).getValue().getUTCDate();
                
                if (currentDay != saveDay || currentMonth != saveMonth) {
                    paraDate = body.appendParagraph(getTextDate(currentSheet.getRange(i, COL_DATE+1).getValue(), false)); 
                    paraDate.setAttributes(STYLE_TITRE1);
                    saveDay = currentDay;
                    saveMonth = currentMonth;
                }

                var paraTypeEvent = body.appendParagraph(currentSheet.getRange(i, COL_TYPE+1).getValue());

                // on constitue le texte avec le nom de l'event et ses horaires & lieux. NB : les Locales ne sont pas gérées par google apps script.                     
                //var heureDeb = currentSheet.getRange(i,COL_HORAIRE_DEB).getValue().toLocaleString("en", {hour12: false, hour: "2-digit", minute: "2-digit"});
                //var heureFin = currentSheet.getRange(i,COL_HORAIRE_FIN).getValue().toLocaleString("en", {hour12: false, hour: "2-digit", minute: "2-digit"});
                var tempDateDeb = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10, 0, 0);                
                var tempDateFin = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10, 0, 0);
                var heureDeb = tempDateDeb.getHours() + "h" + (tempDateDeb.getUTCMinutes()<10?"0"+tempDateDeb.getUTCMinutes() : tempDateDeb.getUTCMinutes());
                var heureFin = tempDateFin.getHours() + "h" + (tempDateFin.getUTCMinutes()<10?"0"+tempDateFin.getUTCMinutes() : tempDateFin.getUTCMinutes());
                var descEvent = currentSheet.getRange(i, COL_TITRE+1).getValue() + "\n" + heureDeb + "-" + heureFin + " - " + currentSheet.getRange(i, COL_LIEU+1).getValue();

                var paraDescEvent = body.appendParagraph(descEvent);

                // on constitue le texte avec les infos de l'event
                var infoEvent = currentSheet.getRange(i, COL_DETAIL+1).getValue();
                var paraInfoEvent = body.appendParagraph(infoEvent);
                
                // saut de ligne
                body.appendParagraph("");
                
                // style des paragraphes
                paraTypeEvent.setAttributes(STYLE_TITRE2);
                paraDescEvent.setAttributes(STYLE_TITRE2);
                paraInfoEvent.setAttributes(STYLE_TEXTE);
            }
        } else if ((currentSheet.getRange(i, COL_CONFIRMATION+1).getValue() == "Confirmé" || currentSheet.getRange(i, COL_CONFIRMATION+1).getValue() == "Diffusé") && currentSheet.getRange(i, COL_PUBLIC+1).getValue() == "buropolis"){ // événements entre buropolis  
            var objCours = {};
            // écrire le titre du cours
            objCours.titre = currentSheet.getRange(i, COL_TITRE+1).getValue();
            
            // écrire la date & heure & lieu
            var dateEvent = getTextDate(currentSheet.getRange(i, COL_DATE+1).getValue(), false);
            var tempDateDeb = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10, 0, 0);                
            var tempDateFin = new Date(2000, 01, 01, currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours(), currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10, 0, 0);
            var heureDeb = tempDateDeb.getHours() + "h" + (tempDateDeb.getUTCMinutes()<10?"0"+tempDateDeb.getUTCMinutes() : tempDateDeb.getUTCMinutes());
            var heureFin = tempDateFin.getHours() + "h" + (tempDateFin.getUTCMinutes()<10?"0"+tempDateFin.getUTCMinutes() : tempDateFin.getUTCMinutes());
            
            // var heureDeb = currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_DEB+1).getValue().getUTCMinutes()+10));
            // var heureFin = currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCHours() + "h" + ((currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10)=="60"?"00":(currentSheet.getRange(i,COL_HORAIRE_FIN+1).getValue().getUTCMinutes()+10));
            objCours.desc = dateEvent + " - " + heureDeb + "-" + heureFin + " - " + currentSheet.getRange(i, COL_LIEU+1).getValue();
            
            // écrire le détail
            objCours.info = currentSheet.getRange(i, COL_DETAIL+1).getValue();               

            // ajout au tableau des cours                
            tabGV.push(objCours);
        }

        // si la date dépasse 7j après la première, on sort de la boucle
        if (currentSheet.getRange(i, COL_DATE+1).getValue() > dateFin) {
            break;
        }
    }

    // on a traité tous les évènements de la semaine, on ajoute les EXPOS & COURS        
    addParagraphFromTab(body, tabExpos);
    addParagraphFromTab(body, tabCoursEtStages);
    // Puis les events Buropolis
    addParagraphFromTab(body, tabGV);
}

/*
 *  fonction renvoyant la date sous forme textuelle à partir d'un objet date
 */
function getTextDate(date, appendYear){
    var text = "";
    switch (date.getDay()) {
        case 0: text+="Dimanche";break;
        case 1: text+="Lundi";break;
        case 2: text+="Mardi";break;
        case 3: text+="Mercredi";break;
        case 4: text+="Jeudi";break;
        case 5: text+="Vendredi";break;
        case 6: text+="Samedi";break;
    }
    text+=" "+date.getDate()+" ";
    switch (date.getMonth()) {
        case 0: text+="janvier";break;
        case 1: text+="février";break;
        case 2: text+="mars";break;
        case 3: text+="avril";break;
        case 4: text+="mai";break;
        case 5: text+="juin";break;
        case 6: text+="juillet";break;
        case 7: text+="août";break;
        case 8: text+="septembre";break;
        case 9: text+="octobre";break;
        case 10: text+="novembre";break;
        case 11: text+="décembre";break;
    }
    if (appendYear){
        text+=" "+date.getFullYear();
    }
    return text;
} 

/*
 *  Fonction ajoutant au body d'un document texte un texte formaté issu d'un tableau de paragraphes
 */
function addParagraphFromTab(body, tabParagraph){
    body.appendPageBreak(); // saut de page        
    var paraTypeEvent = body.appendParagraph(tabParagraph[0]);
    paraTypeEvent.setAttributes(STYLE_TITRE1);
    for (var i = 1 ; i < tabParagraph.length ; i++){
        var paraTitre = body.appendParagraph(tabParagraph[i].titre);        
        var paraDescEvent = body.appendParagraph(tabParagraph[i].desc);
        var paraInfoEvent = body.appendParagraph(tabParagraph[i].info);        
        body.appendParagraph(""); // saut de ligne

        paraTitre.setAttributes(STYLE_TITRE2);
        paraDescEvent.setAttributes(STYLE_TEXTE);
        paraInfoEvent.setAttributes(STYLE_TEXTE);
    }        
}


/*
 * Définition des styles de paragraphes
 */
var STYLE_TITRE1 = {};
STYLE_TITRE1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
STYLE_TITRE1[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
STYLE_TITRE1[DocumentApp.Attribute.FONT_SIZE] = 14;
STYLE_TITRE1[DocumentApp.Attribute.BOLD] = true;

var STYLE_TITRE2 = {};
STYLE_TITRE2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
STYLE_TITRE2[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
STYLE_TITRE2[DocumentApp.Attribute.FONT_SIZE] = 12;
STYLE_TITRE2[DocumentApp.Attribute.BOLD] = true;

var STYLE_TEXTE = {};
STYLE_TEXTE[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
STYLE_TEXTE[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
STYLE_TEXTE[DocumentApp.Attribute.FONT_SIZE] = 8;
STYLE_TEXTE[DocumentApp.Attribute.BOLD] = false;

/*
 * définition des couleurs pour les évènements
 */
var COLORS = {
  PALE_BLUE : 1,
  PALE_GREEN : 2,
  MAUVE : 3,
  PALE_RED : 4,
  YELLOW : 5,
  ORANGE : 6,
  CYAN : 7,
  GRAY : 8,
  BLUE : 9,
  GREEN : 10,
  RED : 11
}

/******************************************************************************************************************************
 *              Utilitaires
 ******************************************************************************************************************************/
/*
 * renvoie l'heure sous forme d'entier depuis une chaîne de caractères formatée de la forme "20h30" ou "20h"
 * si la chaine est juste un nombre, renvoie ce nombre 
 */
function getNbHours(timeStr) {
  if (typeof timeStr == "number") {
    return timeStr;
  }
  if (typeof timeStr == "string") {
    var indexH = timeStr.toLowerCase().indexOf("h");
    if (indexH != -1) {
      return timeStr.substr(0,indexH);
    } else {
      indexH = timeStr.toLowerCase().indexOf(":");
      if (indexH != -1) {
        return timeStr.substr(0,indexH);
      } else {
        return 0;
      }
    }
  }
  if (timeStr instanceof Date) {
    // On ajoute 1h pour coller au fuseau horaire UTC+1
    return timeStr.getUTCHours()+1;
    //return timeStr.getHours();
  }
  return 0;
}  

/*
 * renvoie le nombre de minutes sous forme d'entier depuis une chaîne de caractères formatée de la forme "20h30" ou "20h"
 * si la chaine est juste un nombre, renvoie 0
 */
function getNbMinutes(timeStr) {
  if (typeof timeStr == "string"){
    var indexH = timeStr.toLowerCase().indexOf("h");
    if (indexH != -1) {
      return timeStr.substr(indexH + 1, timeStr.length);
    } else {
      indexH = timeStr.toLowerCase().indexOf(":");
      if (indexH != -1){
        return timeStr.substr(indexH + 1, timeStr.length);
      } else {
        return 0;
      }
    }
  }
  if (timeStr instanceof Date) {
    //return timeStr.getUTCMinutes()+10;
    return timeStr.getUTCMinutes();
  }
  return 0;
}

/*
 *  Fonction de test, pour débug
 */
function testApp(){
      var typeEvent = '✿ JARDIN ✿';
      var color = getColor(typeEvent);
      Logger.log('debug couleur : ' + color);
//    var location = 'Bâtiment Lelong - Amphithéâtre';
//    
//    var calTab = getCalendars(location); // récupère les noms des 3 potentiels calendriers sur lesquels diffuser, tester null
//    
//    if (calTab[0][0] != '') { // diffusion agenda1
//        Logger.log('agenda1 ' + calTab[0][0]);
//        
//        var event = CalendarApp.getOwnedCalendarsByName(calTab[0][0])[0].createEvent('Apollo 11 Landing',
//                     new Date('February 01, 2017 20:00:00 UTC'),
//                     new Date('February 01, 2017 21:00:00 UTC'),
//                     {location: 'The Moon'});
//       Logger.log('Event ID: ' + event.getId());
//
//    }
//
//    if (calTab[0][1] != '') { // diffusion agenda2
//        Logger.log('agenda2 ' + calTab[0][1]);
//    }
//
//    if (calTab[0][2] != '') { // diffusion publique, informations moindre
//        Logger.log('agenda3 ' + calTab[0][2]);
//    }
}
