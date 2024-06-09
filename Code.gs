var SwapperEnums =  {
  INHERIT_TEMPLATE:'INHERIT_TEMPLATE',                         // inherits the attributes/scaling from the template
  KEEP_ORIGINAL:'KEEP_ORIGINAL',                               // keeps the attributes/scaling of the original
  SCALE_ORIGINAL_HEIGHT:'SCALE_ORIGINAL_HEIGHT',               // keeps the height of the original image and scales the width 
  SCALE_ORIGINAL_WIDTH:'SCALE_ORIGINAL_WIDTH'                  // keeps the width of the original image and scales the height 
 
};


function imageSwapper () {

  var params = {
    template:'1rje8nZ-flJTFlBJO3XYm9mzQ-G0mHEkGlCVhEAXydwE',   // this is the document that contains the image swappong template
    scaling:SwapperEnums.SCALE_ORIGINAL_HEIGHT,               // change this as required.                                            
    attributes:SwapperEnums.KEEP_ORIGINAL,
    sections:['Header', 'Footer', 'Body']                     // where to look for images .. normally all 3
  }
  
  // there are twp ways of specifying the documents to look at .. one is simply to provide a list of document ids.
  // as below
  params.documents = [
      '1oI9Pefe5yjkrJKsKw1Mx6ITU_ONdTzZ-PAz3I_d48Qk'          // this is a list of documents to apply image swapping to
  ];
  
  
  // OR another way to populate the documents needed, instead of populating the list manualay
  // specify the starting path of the folder to look ( you can sepcify the id if you prefer)
  // the type of files
  // and whether to recurse (look in lower level paths)
  //
  
 var piles = cUseful.DriveUtils
 .setService(DriveApp)
 .getPileOfFiles (
    "/books/youtube", "application/vnd.google-apps.document", true
  );
  // just interested in the ids.
  params.documents = piles.map (function (d) {
    return d.file.getId();
  });

  
  // do the image swapping
  var result = doImageSwapping ( params );
  
  // if used the piles method, we can enrich the results
  Logger.log (JSON.stringify(result.map (function(d,i) {
    if (piles) {
      d.path = piles[i].path;
      d.folderId = piles[i].folder.getId()
    }
    return d;
  })));
              
}

function doImageSwapping (params) {
  
  // first get the template
  var template = DocumentApp.openById(params.template);
  if (!template) throw 'could not open template document ' + params.template;
  
  var templateImages = getTemplateImages (template.getBody());
  if (!templateImages.length) throw 'no image swaps found in template';
  
  // make sure all the documents exist
  return params.documents.map (function (d) {
    var doc = DocumentApp.openById(d);
    if (!doc) throw 'could not open document ' + d;
    return doc;
  })
  
  // and do the swapping
  .map (function (doc) {
    
    var result = {
      id:doc.getId(),
      name:doc.getName(),
      imagesDetected:0,
      imagesReplaced:0
    };
    
    // make sure we skip the template
    if (result.id === params.template) {
      result.skipped = "Document was same as template - skipped" ;
      return result;
    }
    
    // look in each section asked for
    params.sections.forEach (function(section) {
                  
      // get all the images in the document
      var lump = doc['get'+section]();
      if (lump) {
        lump.getImages().forEach (function(image) {
          
          var sha = cUseful.Utils.blobDigest (image.getBlob());
          result.imagesDetected++;
          
          // check for a match against template
          var match = templateImages.filter (function (d) {
            return sha === d.sha;
          });
          
          // substitute & resize
          if (match.length) {
            
            result.imagesReplaced++;  
            // get the parent of current image, add the new one
            var parent = image.getParent();
            var newImage = parent.insertInlineImage(parent.getChildIndex(image), match[0].replaceWith.getBlob());
            
            // figure out scaling
            var newHeight,newWidth;
            
            if (params.scaling === SwapperEnums.INHERIT_TEMPLATE) {
              newHeight = newImage.getHeight();
              newWidth = newImage.getWidth();
            }
            
            else if (params.scaling === SwapperEnums.KEEP_ORIGINAL) {
              
              newHeight = image.getHeight();
              newWidth = image.getWidth();
            }
            
            else if (params.scaling === SwapperEnums.SCALE_ORIGINAL_HEIGHT) {
              newHeight = image.getHeight();
              newWidth = newImage.getWidth() * newHeight / newImage.getHeight() ;
            }
            
            else if (params.scaling === SwapperEnums.SCALE_ORIGINAL_WIDTH) {
              newWidth = image.getWidth();
              newHeight = newImage.getHeight() * newWidth / newImage.getWidth() ;
            }
            
            else {
              throw  params.scaling + ' is invalid scaling treatment'
            }
            
            // attribute inheritance
            if (params.attributes === SwapperEnums.KEEP_ORIGINAL) {
              newImage.setAttributes(image.getAttributes())
            }
            
            else if (params.attributes !== SwapperEnums.INHERIT_TEMPLATE) {
              throw params.attributes + ' is invalid attribute treatment'
            }
            
            
            // deal with size
            newImage.setHeight (newHeight).setWidth(newWidth);
            
            // delete the old one
            image.removeFromParent();
            
          }
          
        });
      }
    });
    return result;
  });

}



function getTemplateImages(body) {
  
  // return an array of {searchfor:inlineimage, repaceWith:inlineimage , sha:searchSha:the sha of the searched image}
  // table = body.getTables()[0];
  var imagePairs = body.getTables().reduce(function (images,table) {
    getChildrenArray (table).forEach (function (row) {
      var rowImages = [];
      getChildrenArray (row).forEach (function (cell) {
        getChildrenArray (cell).forEach (function (para) {
          getChildrenArray(para.asParagraph()).forEach (function(elem) {
            if (elem.getType() === DocumentApp.ElementType.INLINE_IMAGE) {

              rowImages.push(elem);
            }
          });
        });
      });
      // if exactly two images were found on this row, then they are replacement templates
      if (rowImages.length === 2 ) {
        images.push ({
          searchFor:rowImages[0],
          replaceWith:rowImages[1],
          sha:cUseful.Utils.blobDigest (rowImages[0].getBlob())
        });
      }
    });
    return images;
  },[]);
  
  // check for dups.
  imagePairs.forEach(function (d) {
    if (imagePairs.filter (function (e) {
      return d.sha === e.sha;
    }).length !== 1 ) throw 'duplicate search image in template';
  });
  

  return imagePairs;
  
  function getChildrenArray (parent) {
    var children = [];
    for (var i = 0 ; i < parent.getNumChildren() ; i++) {
      children.push(parent.getChild(i));
    }
    return children;
  }

  
}
