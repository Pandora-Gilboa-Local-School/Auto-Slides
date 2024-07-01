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

function actualizarAjustes(form) {

  // Invocado desde sidePanel_js
  // Al devolver form desde cliente, si una casilla de verificación no está marcada,
  // su propiedad (name) en el objeto pasado a servidor no se devuelve (cuidado).
  
  PropertiesService.getDocumentProperties().setProperties({
    'sAdvance' : form.sAdvance,
    'sReload' : form.sReload,
    'msFade' : form.msFade,
    'backgroundColor' : form.backgroundColor,
    'start' : form.start, // 'on' o NULL
    'repeat' : form.repeat, // 'on' o NULL
    'hideMenu' : form.hideMenu, // 'on' o NULL
    'hideBands' : form.hideBands, // 'on' o NULL
    'hideBorders' : form.hideBorders // 'on' o NULL
  }, false);
  
}

function obtenerRevisiones() {
  
  // Devuelve el ID de la última revisión de la presentación actual
  
  var slideId = SlidesApp.getActivePresentation().getId();
  var respuesta;
  var token;
  var revisiones = [];
  var hayMas = true;
 
  // Iterar hasta alcanzar la última revisión de la presentación
 
  try {
      
    while (hayMas == true) {
      respuesta = Drive.Revisions.list(slideId, {maxResults: 1000, pageToken: token});
      revisiones = revisiones.concat(respuesta.items);
      token = respuesta.nextPageToken;
      hayMas = (token == undefined) ? false : true;
    }
    
    // Devolver última revisión
    
    return revisiones[revisiones.length-1].id;
    
  } catch(e) {
  
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌ Error getting presentation revisions.\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
 
  }
}

function acortarUrl() {
  
  // Invocado desde infoPublicada
  
  var urlCorto = PropertiesService.getDocumentProperties().getProperty('urlCorto');
  
  if (urlCorto == null) {
    
    // No se ha acortado aún, lo haremos ahora y guardaremos URL corto en propiedades
  
    urlCorto = UrlFetchApp.fetch(TINYURL + ScriptApp.getService().getUrl()).getContentText();
    PropertiesService.getDocumentProperties().setProperty('urlCorto', urlCorto);
    
  }
  
  return urlCorto;
  
}
    
function publish() {
     
  var slideId = SlidesApp.getActivePresentation().getId();
  var ultimaRevId = obtenerRevisiones();
  
  // Publicar última revisión de la presentación
 
  try {
      
    Drive.Revisions.patch({published: true,
                           publishedOutsideDomain: true,
                           publishAuto: true}, 
                          slideId, ultimaRevId);
            
    PropertiesService.getDocumentProperties().setProperty('publish', 'true');
    
    // Si no se ha configurado previamente, establecer valores por defecto
    
    if (PropertiesService.getDocumentProperties().getProperty('initialized') != 'true') {
      defaultSettings();
    }    
    
    if (ScriptApp.getService().isEnabled() == true) {
      
      // La webapp ya ha sido previamente publicada, obtener URL público (¡con V8 devuelve el privado /dev! a 18/02/20)
      
      var urlWebApp = ScriptApp.getService().getUrl();
      var panel = HtmlService.createTemplateFromFile('infoPublicada');
            
      panel.url = urlWebApp;
      SlidesApp.getUi().showModalDialog(panel.evaluate().setWidth(700).setHeight(175), '🔄 AutoSlides');
      
    } else {
      
      // El usuario debe realizar la publicación inicial de la webapp

      var panel = HtmlService.createHtmlOutputFromFile('instruccionesWebApp');
      SlidesApp.getUi().showSidebar(panel.setTitle('🌐 Publishing instructions'));

    }
    
    
  } catch(e) {
   
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌ Error publishing presentation.\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
    
  }

}  

function unpublish() {

  var slideId = SlidesApp.getActivePresentation().getId();
  var ultimaRevId = obtenerRevisiones();
 
  // Desactivar publicación de la última revisión de la presentación
 
  try {
  
    Drive.Revisions.patch({published: false,
                         publishedOutsideDomain: false,
                         publishAuto: false}, 
                         slideId, ultimaRevId);
  
    PropertiesService.getDocumentProperties().setProperty('publish', 'false');
    SlidesApp.getUi().alert('🔄 AutoSlides', '🔻 The presentation is no longer publicly available.', SlidesApp.getUi().ButtonSet.OK);
  
  } catch(e) {
   
    SlidesApp.getUi().alert('🔄 AutoSlides', '❌Error stopping publishing presentation\n\n' + e, SlidesApp.getUi().ButtonSet.OK); 
  
  }
  
}

function doGet(e) {

  // Generar página web con presentación incrustada

  var incrustaWeb = HtmlService.createTemplateFromFile('slidesEmbed');
  
  // Rellenar elementos de plantilla
  
  var ajustes = PropertiesService.getDocumentProperties().getProperties();
  var aspecto = 100 * SlidesApp.getActivePresentation().getPageHeight() / SlidesApp.getActivePresentation().getPageWidth();
  var offsetPx = ajustes.hideBorders == 'on' ? BORDER_INSET  : 0;
  
  incrustaWeb.url =  'https://docs.google.com/presentation/d/' + SlidesApp.getActivePresentation().getId() + '/embed';
  incrustaWeb.start = ajustes.start == 'on' ? 'true' : 'false';
  incrustaWeb.repeat = ajustes.repeat == 'on' ? 'true' : 'false';
  incrustaWeb.msAdvance = (+ajustes.sAdvance * 1000).toString();
  incrustaWeb.msFade = ajustes.msFade;
  incrustaWeb.msReload = (+ajustes.sReload * 1000).toString();
  incrustaWeb.backgroundColor = ajustes.backgroundColor;
  incrustaWeb.insetInferior = ajustes.hideMenu == 'on' ? Math.ceil(BOTTOM_INSET  + offsetPx).toString() : '0';
  incrustaWeb.insetLateral = ajustes.hideBands == 'on' ? Math.ceil(100 * MAGIC_NUMBER / aspecto + offsetPx).toString() : '0';
  incrustaWeb.insetSuperior = offsetPx.toString();

  // Para "truco" CSS que hace el iframe responsive

  incrustaWeb.aspecto = aspecto.toString();
  
  return incrustaWeb.evaluate().setTitle(SlidesApp.getActivePresentation().getName()).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}