function createContentsSlide(selectedSlideIndices) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  try {
    if (selectedSlideIndices.length > CONFIG.LIMITS.MAX_TOC_SLIDES) {
      throw new Error(getText("error_too_many_slides"));
    }

    let contentsSlide;
    try {
      const layouts = presentation.getLayouts();
      let titleBodyLayout = null;
      for (let i = 0; i < layouts.length; i++) {
        const layoutName = layouts[i].getLayoutName();
        if (layoutName.includes('TITLE') && layoutName.includes('BODY')) {
          titleBodyLayout = layouts[i];
          break;
        }
      }
      
      if (titleBodyLayout) {
        contentsSlide = presentation.appendSlide(titleBodyLayout);
      } else if (layouts.length > 0) {
        contentsSlide = presentation.appendSlide(layouts[0]);
      } else {
        contentsSlide = presentation.appendSlide();
      }
    } catch (e) {
      contentsSlide = presentation.appendSlide();
    }
    
    const pageElements = contentsSlide.getPageElements();
    let titleElement = null;
    let bodyElement = null;
    
    const pageWidth = presentation.getPageWidth();
    const pageHeight = presentation.getPageHeight();
    
    for (let i = 0; i < pageElements.length; i++) {
      const element = pageElements[i];
      const left = element.getLeft();
      const top = element.getTop();
      const width = element.getWidth();
      const height = element.getHeight();
      
      if (top < pageHeight * 0.2 && width > pageWidth * 0.5) {
        titleElement = element;
      }
      else if (top > pageHeight * 0.2 && width > pageWidth * 0.5 && height > pageHeight * 0.3) {
        bodyElement = element;
      }
    }
    
    if (titleElement && titleElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      try {
        const titleShape = titleElement.asShape();
        if (titleShape.getText()) {
          titleShape.getText().setText(getText("toc_title"));
          titleShape.getText().getTextStyle().setFontSize(32).setBold(true);
          titleShape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      } catch (e) {
        createCustomTitle(contentsSlide, pageWidth);
      }
    } else {
      createCustomTitle(contentsSlide, pageWidth);
    }
    
    let contentArea = null;
    
    if (bodyElement) {
      contentArea = {
        x: bodyElement.getLeft(),
        y: bodyElement.getTop(),
        width: bodyElement.getWidth(),
        height: bodyElement.getHeight()
      };
      try {
        bodyElement.remove();
      } catch (e) {}
    } else {
      contentArea = {
        x: pageWidth * 0.05,
        y: pageHeight * 0.2,
        width: pageWidth * 0.9,
        height: pageHeight - (pageHeight * 0.2) - 50
      };
    }
    
    const totalSlides = selectedSlideIndices.length;
    let column1Slides, column2Slides;
    let useTwoColumns = false;
    
    if (totalSlides > 4) {
      useTwoColumns = true;
      const col1Count = Math.ceil(totalSlides / 2);
      column1Slides = selectedSlideIndices.slice(0, col1Count);
      column2Slides = selectedSlideIndices.slice(col1Count);
    } else {
      column1Slides = selectedSlideIndices;
      column2Slides = [];
    }

    const availableHeight = contentArea.height;
    const maxRows = Math.max(column1Slides.length, column2Slides.length) || 1;
    const ITEM_HEIGHT = Math.min(100, availableHeight / maxRows);
    
    const gap = 15;
    const THUMBNAIL_HEIGHT = Math.max(30, Math.min(68, ITEM_HEIGHT - gap));
    const THUMBNAIL_WIDTH = THUMBNAIL_HEIGHT * (120 / 68);
    const NUMBER_SIZE = Math.max(20, Math.min(30, THUMBNAIL_HEIGHT * 0.5));
    
    const COLUMN1_X = contentArea.x;
    const COLUMN_WIDTH = useTwoColumns ? (contentArea.width - 50) / 2 : contentArea.width;
    const COLUMN2_X = useTwoColumns ? contentArea.x + COLUMN_WIDTH + 50 : 0;
    const START_Y = contentArea.y;
    
    function addSlideToContents(slideIndex, x, y) {
      if (slideIndex >= slides.length) return y;
      const slide = slides[slideIndex];
      const slideNumber = slideIndex + 1;
      const thumbnail = getSlideThumbnail(slide);
      
      if (!thumbnail) return y;
      
      try {
        if (y + ITEM_HEIGHT > pageHeight) return y;
        
        const image = contentsSlide.insertImage(thumbnail, x, y, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT);
        try { image.setLinkSlide(slide); } catch (e) {}
        
        const numberBox = contentsSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, NUMBER_SIZE, NUMBER_SIZE);
        numberBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
        numberBox.getBorder().setTransparent();
        try { numberBox.setLinkSlide(slide); } catch (e) {}
        
        const numberText = contentsSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, NUMBER_SIZE, NUMBER_SIZE);
        numberText.getText().setText(slideNumber.toString());
        numberText.getText().getTextStyle().setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
        numberText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        try { numberText.setLinkSlide(slide); } catch (e) {}
        
        const descriptionX = x + THUMBNAIL_WIDTH + 15;
        const descriptionY = y + 10;
        const descriptionWidth = Math.min(COLUMN_WIDTH - THUMBNAIL_WIDTH - 20, 400);
        const descriptionHeight = 50;
        
        const description = contentsSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, descriptionX, descriptionY, descriptionWidth, descriptionHeight);
        const slideTitle = getSlideTitle(slide, slideIndex);
        description.getText().setText(slideTitle);
        description.getText().getTextStyle().setFontSize(14);
        try { description.setLinkSlide(slide); } catch (e) {}
        
        return y + ITEM_HEIGHT;
      } catch (e) {
        return y;
      }
    }
    
    let currentY = START_Y;
    for (const slideIndex of column1Slides) {
      currentY = addSlideToContents(slideIndex, COLUMN1_X, currentY);
    }
    
    if (useTwoColumns && column2Slides.length > 0) {
      currentY = START_Y;
      for (const slideIndex of column2Slides) {
        currentY = addSlideToContents(slideIndex, COLUMN2_X, currentY);
      }
    }
    
    addSpeakerNotesHint(contentsSlide, totalSlides);
    contentsSlide.selectAsCurrentPage();
    return contentsSlide; 
  } catch (error) {
    throw new Error(error.toString());
  }
}

function addSpeakerNotesHint(slide, slidesCount) {
  try {
    const notesPage = slide.getNotesPage();
    let notesShape = notesPage.getSpeakerNotesShape();
    let hintStr = getText("notes_hint");
    hintStr = hintStr.replace('%s', slidesCount).replace('%s', new Date().toLocaleString());

    if (notesShape) {
      notesShape.getText().setText(hintStr);
    } else {
      const pres = SlidesApp.getActivePresentation();
      notesPage.insertTextBox(hintStr, 50, 50, pres.getPageWidth() - 100, pres.getPageHeight() - 100);
    }
  } catch (e) {}
}

function createCustomTitle(slide, pageWidth) {
  const title = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, pageWidth * 0.05, 50, pageWidth * 0.9, 60);
  title.getText().setText(getText("toc_title"));
  title.getText().getTextStyle().setFontSize(32).setBold(true);
  title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  return title;
}

function addNavigationToSlides(contentsSlide) {
  const presentation = SlidesApp.getActivePresentation();
  const allSlides = presentation.getSlides();
  const totalSlides = allSlides.length;

  if (!contentsSlide) contentsSlide = findContentsSlide();
  if (!contentsSlide) throw new Error(getText("error_no_toc_slide"));

  let slidesToUpdate = [];
  for (let i = 1; i < allSlides.length; i++) {
    if (allSlides[i].getObjectId() !== contentsSlide.getObjectId()) {
      slidesToUpdate.push(allSlides[i]);
    }
  }

  let count = 0;
  const pageWidth = presentation.getPageWidth();
  const pageHeight = presentation.getPageHeight();
  
  const footerHeight = 35;
  const footerWidth = 60;
  const paddingRight = 15;
  const paddingBottom = 15;
  const footerX = pageWidth - footerWidth - paddingRight;
  const footerY = pageHeight - footerHeight - paddingBottom;

  slidesToUpdate.forEach(slide => {
    removeExistingNavigation(slide);

    const bg = slide.insertShape(SlidesApp.ShapeType.FOLDED_CORNER, footerX, footerY, footerWidth, footerHeight);
    bg.getFill().setSolidFill(CONFIG.COLORS.NAV_BG);
    bg.getBorder().setTransparent();
    bg.setTitle(CONFIG.ELEMENTS.NAV_TITLE);
    
    const currentSlideNum = allSlides.indexOf(slide) + 1;
    const textStr = `${currentSlideNum} / ${totalSlides}`;

    const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, footerX, footerY, footerWidth, footerHeight);
    textBox.getFill().setTransparent();
    textBox.getBorder().setTransparent();
    textBox.setTitle(CONFIG.ELEMENTS.NAV_TEXT_TITLE);
    
    const txtRange = textBox.getText();
    txtRange.setText(textStr);
    txtRange.getTextStyle().setFontSize(11).setForegroundColor(CONFIG.COLORS.NAV_TEXT).setBold(true);
    txtRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    try {
      bg.setLinkSlide(contentsSlide);
      textBox.setLinkSlide(contentsSlide);
    } catch (e) {}
    count++;
  });
  return count;
}

function removeExistingNavigation(slide) {
  const elements = slide.getPageElements();
  const elementsToRemove = [];

  elements.forEach(el => {
    if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      try {
        const shape = el.asShape();
        const title = shape.getTitle();
        if (title === CONFIG.ELEMENTS.NAV_TITLE || title === CONFIG.ELEMENTS.NAV_TEXT_TITLE) {
          elementsToRemove.push(el);
        }
      } catch (e) {}
    }
  });

  elementsToRemove.forEach(el => {
    try { el.remove(); } catch (e) {}
  });
}

function quickCreate() {
  try {
    const ui = SlidesApp.getUi();
    const slides = SlidesApp.getActivePresentation().getSlides();
    if (slides.length <= 1) {
      ui.alert(getText("sidebar_title"), getText("error_not_enough_slides"), ui.ButtonSet.OK);
      return;
    }
    const selectedSlideIndices = [];
    for (let i = 1; i < slides.length; i++) selectedSlideIndices.push(i);
    
    const response = ui.alert(getText("quick_create_title"), getText("quick_create_prompt"), ui.ButtonSet.OK_CANCEL);
    if (response == ui.ButtonSet.CANCEL) return;
    
    const contentsSlide = createContentsSlide(selectedSlideIndices);
    if (contentsSlide) addNavigationToSlides(contentsSlide);
    
    ui.alert(getText("quick_create_success_title"), getText("quick_create_success_msg"), ui.ButtonSet.OK);
  } catch (error) {
    SlidesApp.getUi().alert('Error', error.toString(), SlidesApp.getUi().ButtonSet.OK);
  }
}

function addNavigationOnly() {
  try {
    const ui = SlidesApp.getUi();
    const slides = SlidesApp.getActivePresentation().getSlides();
    if (slides.length <= 1) {
      ui.alert(getText("sidebar_title"), getText("error_not_enough_slides"), ui.ButtonSet.OK);
      return;
    }
    const contentsSlide = findContentsSlide();
    if (!contentsSlide) {
      ui.alert(getText("sidebar_title"), getText("error_no_toc_slide"), ui.ButtonSet.OK);
      return;
    }
    const count = addNavigationToSlides(contentsSlide);
    let msg = getText("add_nav_success_msg").replace('%s', count);
    ui.alert(getText("add_nav_success_title"), msg, ui.ButtonSet.OK);
  } catch (error) {
    SlidesApp.getUi().alert('Error', error.toString(), SlidesApp.getUi().ButtonSet.OK);
  }
}

function findContentsSlide() {
  try {
    const slides = SlidesApp.getActivePresentation().getSlides();
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const elements = slide.getPageElements();
      for (let j = 0; j < elements.length; j++) {
        const element = elements[j];
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          try {
            const text = element.asShape().getText().asString().trim();
            const titleVariants = ['Содержание', 'Contents', 'Оглавление', getText("toc_title")];
            if (titleVariants.includes(text)) return slide;
          } catch (e) {}
        }
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

function deleteNavigation() {
  try {
    const ui = SlidesApp.getUi();
    const slides = SlidesApp.getActivePresentation().getSlides();
    const response = ui.alert(getText("delete_nav_title"), getText("delete_nav_prompt"), ui.ButtonSet.YES_NO);
    if (response != ui.Button.YES) return;
    
    let totalDeleted = 0;
    slides.forEach(slide => {
      const beforeCount = slide.getPageElements().length;
      removeExistingNavigation(slide);
      const afterCount = slide.getPageElements().length;
      totalDeleted += (beforeCount - afterCount);
    });
    
    let msg = getText("delete_nav_success_msg").replace('%s', totalDeleted).replace('%s', slides.length);
    ui.alert(getText("delete_nav_success_title"), msg, ui.ButtonSet.OK);
  } catch (error) {
    SlidesApp.getUi().alert('Error', error.toString(), SlidesApp.getUi().ButtonSet.OK);
  }
}

function getSlideThumbnail(slide) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const thumbnail = Slides.Presentations.Pages.getThumbnail(
      presentation.getId(),
      slide.getObjectId(),
      { "thumbnailProperties.mimeType": "PNG", "thumbnailProperties.thumbnailSize": "MEDIUM" }
    );
    return thumbnail.contentUrl;
  } catch (error) {
    return null;
  }
}

function getSlideTitle(slide, index) {
  try {
    const shapes = slide.getShapes().filter(s => s.getText() && s.getText().asString().trim().length > 0);
    if (shapes.length === 0) return "Slide " + (index + 1);
    const title = shapes[0].getText().asString().trim();
    return title.length > 50 ? title.substring(0, 47) + "..." : title;
  } catch (error) {
    return "Slide " + (index + 1);
  }
}

function getSlideBasicInfo() {
  const slidesList = SlidesApp.getActivePresentation().getSlides();
  let max = slidesList.length;
  if (max > CONFIG.LIMITS.MAX_SLIDES_FETCH) {
    max = CONFIG.LIMITS.MAX_SLIDES_FETCH;
  }
  
  let result = [];
  for (let i = 0; i < max; i++) {
    const slide = slidesList[i];
    result.push({
      index: i,
      number: i + 1,
      id: slide.getObjectId(),
      title: getSlideTitle(slide, i)
    });
  }
  return result;
}

function getSlideThumbnailById(slideId) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const slide = presentation.getSlideById(slideId);
    if (!slide) return null;
    return getSlideThumbnail(slide);
  } catch (e) {
    return null;
  }
}