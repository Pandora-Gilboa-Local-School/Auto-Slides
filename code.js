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

var INSET_INFERIOR = 28; // Altura en px barra con botones inferior en presentación incrustada
var NUMERO_MAGICO = 14.25; // x ancho/alto para obtener el valor en px del recorte lateral que elimina bandas negras
var INSET_BORDES = 2;      // Desplazamiento adicional para eliminar todos los bordes mediante clip-path / inset (CSS) 
var TINYURL = 'https://tinyurl.com/api-create.php?url='; // URL para acortar usando servicio de TinyURL  

// Vamos a por faena...

function onOpen() {

  SlidesApp.getUi().createMenu('🔄 AutoSlides')
    .addItem('⚙️ Configure', 'configurar')
    .addItem('🌐  Get Public URL', 'publicar')
    .addItem('🔻 Stop Publishing', 'despublicar')
    .addSeparator()
    .addItem('💡 About AutoSlides', 'acercaDe')
    .addToUi();
    
}

// Info del script

function acercaDe() {

  // Presentación del complemento
  var panel = HtmlService.createTemplateFromFile('acercaDe');
  panel.version = VERSION;
  SlidesApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(375), '💡 What is autoslides?');
}

// Refrescar gráficos vinculados de HdC

function refrescarGraficosHdc() {
    
  // Versión V8: No se utiliza por bug V8 y ScriptApp.GetService().getUrl()
  // https://groups.google.com/d/topic/google-apps-script-community/0snPFcUqt40/discussion
  
  // SlidesApp.getActivePresentation().getSlides().map(diapo => {diapo.getSheetsCharts().map(grafico => {grafico.refresh();});});
  
  SlidesApp.getActivePresentation().getSlides().map(function(diapo) {
    diapo.getSheetsCharts().map(function(grafico) {
      grafico.refresh();});});
  
}

function contarGraficosHdc() {
  
  var numGraficos = 0;
  
  // Idem anterior: Versión V8. No se utiliza para seguir ejecutando con Rhino por bug V8 y ScriptApp.GetService().getUrl()

  // SlidesApp.getActivePresentation().getSlides().map(diapo => {diapo.getSheetsCharts().map(grafico => {numGraficos++});});
  
  SlidesApp.getActivePresentation().getSlides().map(function(diapo)
    {diapo.getSheetsCharts().map(function(grafico) 
      {numGraficos++});});
  
  return numGraficos;

}

function configurar() {

  // Inicializar y / o leer configuración
  
  if (PropertiesService.getDocumentProperties().getProperty('initialized') != 'true') {
    
    // Establecer ajustes por defecto
    
    PropertiesService.getDocumentProperties().setProperties(SETTINGS, true);
    
    // Inicialmente la publicación está desactivada
    
    PropertiesService.getDocumentProperties().setProperty('publicar', 'false');
    
  }
  
  // Plantilla del panel
  
  var panel = HtmlService.createTemplateFromFile('panelLateral');
  
  // Valores iniciales de controles
  
  var ajustes = PropertiesService.getDocumentProperties();
  
  panel.sAdvance =  ajustes.getProperty('sAdvance');
  panel.sReload = ajustes.getProperty('sReload');
  panel.msFade = ajustes.getProperty('msFade');
  panel.backgroundColor = ajustes.getProperty('backgroundColor');
  panel.start =  ajustes.getProperty('start') == 'on' ? 'checked' : '' ;
  panel.repeat =  ajustes.getProperty('repeat') == 'on' ? 'checked' : '' ;
  panel.hideMenu = ajustes.getProperty('hideMenu')  == 'on' ? 'checked' : '';
  panel.hideBands = ajustes.getProperty('hideBands')  == 'on' ? 'checked' : '';
  panel.hideBorders = ajustes.getProperty('hideBorders')  == 'on' ? 'checked' : '';
  panel.numGraficos = contarGraficosHdc();
  
  // Construir y desplegar panel de configuración
  
  SlidesApp.getUi().showSidebar(panel.evaluate().setTitle('🔄 AutoSlides: Embed Settings'));
  
}

function ajustesPorDefecto() {
  
  // Invocado desde panelLateral_js
  // Restablecer ajustes por defecto (,false para preservar otras propiedades)
  
  PropertiesService.getDocumentProperties().setProperties(SETTINGS, false);
  
  // Devolver a panelLateral_js para que actualice formulario
  return SETTINGS;
  
}

function actualizarAjustes(form) {

  // Invocado desde panelLateral_js
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
    
function publicar() {
     
  var slideId = SlidesApp.getActivePresentation().getId();
  var ultimaRevId = obtenerRevisiones();
  
  // Publicar última revisión de la presentación
 
  try {
      
    Drive.Revisions.patch({published: true,
                           publishedOutsideDomain: true,
                           publishAuto: true}, 
                          slideId, ultimaRevId);
            
    PropertiesService.getDocumentProperties().setProperty('publicar', 'true');
    
    // Si no se ha configurado previamente, establecer valores por defecto
    
    if (PropertiesService.getDocumentProperties().getProperty('initialized') != 'true') {
      ajustesPorDefecto();
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

function despublicar() {

  var slideId = SlidesApp.getActivePresentation().getId();
  var ultimaRevId = obtenerRevisiones();
 
  // Desactivar publicación de la última revisión de la presentación
 
  try {
  
    Drive.Revisions.patch({published: false,
                         publishedOutsideDomain: false,
                         publishAuto: false}, 
                         slideId, ultimaRevId);
  
    PropertiesService.getDocumentProperties().setProperty('publicar', 'false');
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
  var offsetPx = ajustes.hideBorders == 'on' ? INSET_BORDES : 0;
  
  incrustaWeb.url =  'https://docs.google.com/presentation/d/' + SlidesApp.getActivePresentation().getId() + '/embed';
  incrustaWeb.start = ajustes.start == 'on' ? 'true' : 'false';
  incrustaWeb.repeat = ajustes.repeat == 'on' ? 'true' : 'false';
  incrustaWeb.msAdvance = (+ajustes.sAdvance * 1000).toString();
  incrustaWeb.msFade = ajustes.msFade;
  incrustaWeb.msReload = (+ajustes.sReload * 1000).toString();
  incrustaWeb.backgroundColor = ajustes.backgroundColor;
  incrustaWeb.insetInferior = ajustes.hideMenu == 'on' ? Math.ceil(INSET_INFERIOR  + offsetPx).toString() : '0';
  incrustaWeb.insetLateral = ajustes.hideBands == 'on' ? Math.ceil(100 * NUMERO_MAGICO / aspecto + offsetPx).toString() : '0';
  incrustaWeb.insetSuperior = offsetPx.toString();

  // Para "truco" CSS que hace el iframe responsive

  incrustaWeb.aspecto = aspecto.toString();
  
  return incrustaWeb.evaluate().setTitle(SlidesApp.getActivePresentation().getName()).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}