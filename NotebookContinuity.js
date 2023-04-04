let documentPrp = PropertiesService.getDocumentProperties();

function onOpen() {
  createMenu();
}

function storeContFrom() {
  storeCont("contFrom");
}


function storeContOn() {
  storeCont("contOn");
}

function createMenu() {
  const ui = SlidesApp.getUi();
  ui.createMenu('Continued Headers')
      .addItem('Store Continued-From', 'storeContFrom')
      .addItem('Store Continued-On', 'storeContOn')
      .addSeparator()
      .addItem('Continufy Page Selections', 'labelPageRange')
      .addToUi();
}

function storeCont(contType) {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();

  console.log(contType)
  console.log(JSON.stringify(selection))
  console.log(JSON.stringify(selectionType))

  if (selectionType == SlidesApp.SelectionType.PAGE_ELEMENT) {
     const pageElements = selection.getPageElementRange().getPageElements();
     if (pageElements.length == 1) {
       const element = pageElements[0];
       const elementText = element.asShape().getText();

       const left = element.getLeft();
       const top = element.getTop();
       const width = element.getWidth();
       const height = element.getHeight();
       const font = elementText.getTextStyle().getFontFamily();
       const fontSize = elementText.getTextStyle().getFontSize();

       const colorType = elementText.getTextStyle().getForegroundColor().getColorType();

       if (colorType == SlidesApp.ColorType.RGB) {
         var color = elementText.getTextStyle().getForegroundColor().asRgbColor().asHexString();
       } else if (colorType == SlidesApp.ColorType.THEME) {
         var color = elementText.getTextStyle().getForegroundColor().asThemeColor().getThemeColorType();
       } else {
         var color = null
       }

       console.log(left, top, width, height);
       documentPrp.setProperty((contType + 'Style'), JSON.stringify({
         'left': left,
         'top': top,
         'width': width,
         'height': height,
         'font': font,
         'fontSize': fontSize,
         'colorType': colorType,
         'color': color,
         }));

       console.log(JSON.parse(documentPrp.getProperty((contType + 'Style'))));
     }
  }
}

function labelPageRange() {
  const slides = SlidesApp.getActivePresentation();
  const selection = slides.getSelection();
  const pageRange = selection.getPageRange();

  if (pageRange != null) {
    console.log("entered")
    const pages = pageRange.getPages();
    console.log(pages);
    const firstPageId = pages[0].getObjectId();
    console.log(firstPageId);
    const firstPageNumber = getPageNumber(firstPageId); 
    console.log(firstPageNumber);
    if (pages.length > 0) {
      for (let i = 0; i < pages.length; i++) {
        const currentPageNumber = firstPageNumber + i

        console.log(JSON.parse(documentPrp.getProperty('contFromStyle')));
        console.log(JSON.parse(documentPrp.getProperty('contOnStyle')));

        let parsedContFrom = JSON.parse(documentPrp.getProperty('contFromStyle'));
        let parsedContOn = JSON.parse(documentPrp.getProperty('contOnStyle'));

        let contFromObj = pages[i].insertShape(
          SlidesApp.ShapeType.TEXT_BOX, 
          parsedContFrom['left'], 
          parsedContFrom['top'], 
          parsedContFrom['width'], 
          parsedContFrom['height']
          );
        let contOnObj = pages[i].insertShape(
          SlidesApp.ShapeType.TEXT_BOX,
          parsedContOn['left'],
          parsedContOn['top'],
          parsedContOn['width'],
          parsedContOn['height']
          );

        if (pages.length == 1) {
          contFromObj.getText().setText("(Continued from page —)");
          contOnObj.getText().setText("(Continued on page —)");
        } else if (i == 0) {
          contFromObj.getText().setText("(Continued from page —)");
          contOnObj.getText().setText("(Continued on page " + (currentPageNumber + 1).toString() + ")");
        } else if (i == (pages.length - 1)) {
          contFromObj.getText().setText("(Continued from page " + (currentPageNumber - 1).toString() + ")");
          contOnObj.getText().setText("(Continued on page —)");
        } else {
          contFromObj.getText().setText("(Continued from page " + (currentPageNumber - 1).toString() + ")")
          contOnObj.getText().setText("(Continued on page " + (currentPageNumber + 1).toString() + ")");
        }

        if (parsedContFrom['font'] != null && parsedContFrom['fontSize'] != null) {
          contFromObj.getText().getTextStyle().setFontFamily(parsedContFrom['font']);
          contFromObj.getText().getTextStyle().setFontSize(parsedContFrom['fontSize']);
        }
        if (parsedContOn['font'] != null && parsedContOn['fontSize'] != null) {
          contOnObj.getText().getTextStyle().setFontFamily(parsedContOn['font']);
          contOnObj.getText().getTextStyle().setFontSize(parsedContOn['fontSize']);
        }  

        if (parsedContFrom['colorType'] == 'RGB') {
          contFromObj.getText().getTextStyle().setForegroundColor(parsedContFrom['color']);
        } else if (parsedContFrom['colorType'] == 'THEME') {
          contFromObj.getText().getTextStyle().setForegroundColor( SlidesApp.ThemeColorType[parsedContFrom['color']] );
        }

        if (parsedContOn['colorType'] == 'RGB') {
          contOnObj.getText().getTextStyle().setForegroundColor(parsedContFrom['color']);
        } else if (parsedContOn['colorType'] == 'THEME') {
          contOnObj.getText().getTextStyle().setForegroundColor( SlidesApp.ThemeColorType[parsedContFrom['color']] );
        }
      }
    }
  }
}


function getPageNumber(page_id) {
  const pages = SlidesApp.getActivePresentation().getSlides();
  for(let i = 0; i < pages.length; i++) {
    if(pages[i].getObjectId() == page_id) {
      return i+1;
    }
  }
}

