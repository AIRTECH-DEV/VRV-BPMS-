function doGet() {
  return HtmlService.createHtmlOutputFromFile('index') // Replace 'index' with your actual filename
      .setTitle('Test Login Page')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}