/**
 * AutoSlides
 * A Google Slides template that allows for the creation of a self-hosted slideshow
 * via a web app that updates automatically at preset intervals without the need
 * for manual browser reloading.
 *
 * Copyright (C) 
 *
 * @OnlyCurrentDoc
 */

// General constants for the script

var VERSION = 'Version: 1';

var SETTINGS = {
  'initialized' : 'true',
  'sAdvance' : '3',
  'sReload' : '60',
  'msFade' : '1500',
  'backgroundColor' : '#ffffff',
  'start' : 'on',
  'repeat' : 'on',
  'hideMenu' : 'on',
  'hideBands' : 'on',
  'hideBorders' : 'off'};

  var BOTTOM_INSET = 28; // Height in pixels of the bottom bar with buttons in embedded presentation
  var MAGIC_NUMBER = 14.25; // Width/height ratio to obtain the pixel value of the lateral crop that removes black bands
  var BORDER_INSET = 2; // Additional offset to remove all borders using clip-path / inset (CSS)
  var TINYURL = 'https://tinyurl.com/api-create.php?url='; // URL for shortening using TinyURL service

// Let's get to work...

function onOpen() {

  SlidesApp.getUi().createMenu('🔄 AutoSlides')
    .addItem('⚙️ Configure', 'configure')
    .addItem('🌐  Get Public URL', 'publish')
    .addItem('🔻 Stop Publishing', 'unpublish')
    .addSeparator()
    .addItem('💡 About AutoSlides', 'about')
    .addToUi();
    
}

// Script Info

function about() {
  // Plugin presentation
  var panel = HtmlService.createTemplateFromFile('about');
  panel.version = VERSION;
  SlidesApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(375), '💡 What is autoslides?');
}


// Refresh linked charts from sheets

function refreshHdCCharts() {
    
  // V8 version: Not used due to V8 bug and ScriptApp.GetService().getUrl()
  // https://groups.google.com/d/topic/google-apps-script-community/0snPFcUqt40/discussion
  
  // SlidesApp.getActivePresentation().getSlides().map(slide => {slide.getSheetsCharts().map(chart => {chart.refresh();});});
  
  SlidesApp.getActivePresentation().getSlides().map(function(slide) {
    slide.getSheetsCharts().map(function(chart) {
      chart.refresh();
    });
  });
}

function countSheetCharts() {
  
  var numCharts = 0;
  
  // Same as above: V8 Version. Not used to continue running with Rhino due to V8 bug and ScriptApp.GetService().getUrl()

  // SlidesApp.getActivePresentation().getSlides().map(slide => {slide.getSheetsCharts().map(chart => {numCharts++});});
  
  SlidesApp.getActivePresentation().getSlides().map(function(slide) {
    slide.getSheetsCharts().map(function(chart) {
      numCharts++;
    });
  });
  
  return numCharts;
}




function configure() {

  // Initialize and/or read configuration
  
  if (PropertiesService.getDocumentProperties().getProperty('initialized') != 'true') {
    
    // Set default settings
    
    PropertiesService.getDocumentProperties().setProperties(SETTINGS, true);
    
    // Initially, publishing is disabled
    
    PropertiesService.getDocumentProperties().setProperty('publish', 'false');
    
  }
  
  // Panel template
  
  var panel = HtmlService.createTemplateFromFile('sidePanel');
  
  // Initial control values
  
  var settings = PropertiesService.getDocumentProperties();
  
  panel.sAdvance =  settings.getProperty('sAdvance');
  panel.sReload = settings.getProperty('sReload');
  panel.msFade = settings.getProperty('msFade');
  panel.backgroundColor = settings.getProperty('backgroundColor');
  panel.start =  settings.getProperty('start') == 'on' ? 'checked' : '' ;
  panel.repeat =  settings.getProperty('repeat') == 'on' ? 'checked' : '' ;
  panel.hideMenu = settings.getProperty('hideMenu') == 'on' ? 'checked' : '';
  panel.hideBands = settings.getProperty('hideBands') == 'on' ? 'checked' : '';
  panel.hideBorders = settings.getProperty('hideBorders') == 'on' ? 'checked' : '';
  panel.numCharts = countSheetCharts();
  
  // Build and display configuration panel
  
  SlidesApp.getUi().showSidebar(panel.evaluate().setTitle('🔄 AutoSlides: Embed Settings'));
}




function defaultSettings() {
  
  // Invoked from sidePanel_js
  // Reset to default settings (false to preserve other properties)
  
  PropertiesService.getDocumentProperties().setProperties(SETTINGS, false);
  
  // Return to sidePanel_js to update the form
  return SETTINGS;
}

function updateSettings(form) {

  // Invoked from sidePanel_js
  // When returning the form from the client, if a checkbox is not checked,
  // its property (name) in the object passed to the server is not returned (be careful).
  
  PropertiesService.getDocumentProperties().setProperties({
    'sAdvance' : form.sAdvance,
    'sReload' : form.sReload,
    'msFade' : form.msFade,
    'backgroundColor' : form.backgroundColor,
    'start' : form.start, // 'on' or NULL
    'repeat' : form.repeat, // 'on' or NULL
    'hideMenu' : form.hideMenu, // 'on' or NULL
    'hideBands' : form.hideBands, // 'on' or NULL
    'hideBorders' : form.hideBorders // 'on' or NULL
  }, false);
  
}

function getRevisions() {
  
  // Returns the ID of the latest revision of the current presentation
  
  var slideId = SlidesApp.getActivePresentation().getId();
  var response;
  var token;
  var revisions = [];
  var more = true;
 
  // Iterate until reaching the last revision of the presentation
 
  try {
      
    while (more == true) {
      response = Drive.Revisions.list(slideId, {maxResults: 1000, pageToken: token});
      revisions = revisions.concat(response.items);
      token = response.nextPageToken;
      more = (token == undefined) ? false : true;
    }
    
    // Return the last revision
    
    return revisions[revisions.length - 1].id;
    
  } catch(e) {
  
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌ Error getting presentation revisions.\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
 
  }
}

function shortenUrl() { // Stupid Function.... Just gonna make everything null/blank because I CAN
  
  // Invoked from publishedInfo
  
  var shortUrl = null
  
  if (shortUrl == null) {
    
    // Not yet shortened, let's do it now and store short URL in properties
    
    shortUrl = "" //UrlFetchApp.fetch(TINYURL + ScriptApp.getService().getUrl()).getContentText();
    PropertiesService.getDocumentProperties().setProperty('shortUrl', shortUrl);
    
  }
  
  return shortUrl;
  
}
    
function publish() {
     
  var slideId = SlidesApp.getActivePresentation().getId();
  var lastRevId = getRevisions();
  
  // Publish the latest revision of the presentation
 
  try {
      
    Drive.Revisions.patch({published: true,
                           publishedOutsideDomain: true,
                           publishAuto: true}, 
                          slideId, lastRevId);
            
    PropertiesService.getDocumentProperties().setProperty('publish', 'true');
    
    // If not previously configured, set default values
    
    if (PropertiesService.getDocumentProperties().getProperty('initialized') != 'true') {
      defaultSettings();
    }    
    
    if (ScriptApp.getService().isEnabled() == true) {
      
      // The web app has been previously published, get public URL (V8 returns private /dev as of 18/02/20)
      function getUrl(){
        google.script.run
        .withSuccessHandler(function(url){
           document.getElementById('marco2').innerHTML='<p>' + url + '</p>';
         })
        .getScriptUrl()
      }

      try {
        var webAppUrl = getUrl(); // Assuming getUrl() is your function to fetch URL
        // ScriptApp.getService().getUrl() can be used alternatively if needed
      
      } catch (error) {
        // Handle the error appropriately
        console.error("Error fetching web app URL:", error);
        SlidesApp.getUi().alert("Failed to fetch web app URL. Please check deployment settings.\n\n (Look... if you got this error... I can't help you, I've spent too long trying to debug this stupid function... If you wanna try.. good luck!)");
        SlidesApp.getUi().alert("Your WebApp is now public though, but you need to go through app Scripts to find the URL (hit close to see more info)");

        // You can also add a more descriptive message or log additional information
        console.warn("If you see this error, please verify your deployment settings and script configuration.");
        // Instructions for the user...
        var panel = HtmlService.createHtmlOutputFromFile('findingWebAppUrlInstructions');
        SlidesApp.getUi().showSidebar(panel.setTitle('🌐 Finding the Web App URL'));
      }
      var panel = HtmlService.createTemplateFromFile('publishedInfo');
            
      panel.url = webAppUrl;
      SlidesApp.getUi().showModalDialog(panel.evaluate().setWidth(700).setHeight(175), '🔄 AutoSlides');
      
    } else {
      
      // User needs to perform initial web app publishing

      var panel = HtmlService.createHtmlOutputFromFile('webAppInstructions');
      SlidesApp.getUi().showSidebar(panel.setTitle('🌐 Publishing instructions'));

    }
    
    
  } catch(e) {
   
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌ Error publishing presentation.\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
    
  }

}

function unpublish() {

  var slideId = SlidesApp.getActivePresentation().getId();
  var lastRevId = getRevisions();
 
  // Disable publishing of the latest revision of the presentation
 
  try {
  
    Drive.Revisions.patch({published: false,
                           publishedOutsideDomain: false,
                           publishAuto: false}, 
                          slideId, lastRevId);
  
    PropertiesService.getDocumentProperties().setProperty('publish', 'false');
    SlidesApp.getUi().alert('🔄 AutoSlides', '🔻 The presentation is no longer publicly available.', SlidesApp.getUi().ButtonSet.OK);
  
  } catch(e) {
   
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌Error stopping publishing presentation\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
  
  }
  
}

function doGet(e) {

  // Generate web page with embedded presentation

  var embedWebPage = HtmlService.createTemplateFromFile('slidesEmbed');
  
  // Fill template elements
  
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var aspectRatio = 100 * SlidesApp.getActivePresentation().getPageHeight() / SlidesApp.getActivePresentation().getPageWidth();
  var offsetPx = settings.hideBorders == 'on' ? BORDER_INSET  : 0;
  
  embedWebPage.url =  'https://docs.google.com/presentation/d/' + SlidesApp.getActivePresentation().getId() + '/embed';
  embedWebPage.start = settings.start == 'on' ? 'true' : 'false';
  embedWebPage.repeat = settings.repeat == 'on' ? 'true' : 'false';
  embedWebPage.msAdvance = (+settings.sAdvance * 1000).toString();
  embedWebPage.msFade = settings.msFade;
  embedWebPage.msReload = (+settings.sReload * 1000).toString();
  embedWebPage.backgroundColor = settings.backgroundColor;
  embedWebPage.bottomInset = settings.hideMenu == 'on' ? Math.ceil(BOTTOM_INSET  + offsetPx).toString() : '0';
  embedWebPage.sideInset = settings.hideBands == 'on' ? Math.ceil(100 * MAGIC_NUMBER / aspectRatio + offsetPx).toString() : '0';
  embedWebPage.topInset = offsetPx.toString();

  // For CSS trick making the iframe responsive

  embedWebPage.aspectRatio = aspectRatio.toString();
  
  return embedWebPage.evaluate().setTitle(SlidesApp.getActivePresentation().getName()).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}