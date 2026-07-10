function onOpen() {
  SlidesApp.getUi()
    .createMenu(getText("sidebar_title"))
    .addItem(getText("menu_open_sidebar"), "openSidebar")
    .addItem(getText("menu_quick_create"), "quickCreate")
    .addSeparator()
    .addItem(getText("menu_add_nav"), "addNavigationOnly")
    .addItem(getText("menu_delete_nav"), "deleteNavigation")
    .addSeparator()
    .addItem(getText("menu_help"), "showHelp")
    .addToUi();
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function openSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar").evaluate()
      .setTitle(getText("sidebar_title"));
  SlidesApp.getUi().showSidebar(html);
}

function showHelp() {
  const htmlContent = getText("help_html");
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(300);
  
  SlidesApp.getUi().showModalDialog(html, 'Help');
}
