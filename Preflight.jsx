
// Main Code [Execution of script begins here]

// uncomment to suppress Illustrator warning dialogs
// app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

// Turn debugging mesages on or off
var debugThisScript = false;
var xdebugThisScript = false;

// Which version of this script
// 0.1 First version
// 0.11 Minor changes here and there.
// 0.12 Wrap fill color checking in a try/catch block and flag type with crazy CMYK values.
var scriptVersion = "0.12";

// find active document
var theDoc = app.activeDocument;

// Arrays to hold details on the problem elements we find.
var badBlackType = new Array();
var badBlackIndex = 0;

var badOverprintType = new Array();
var badOverprintIndex = 0;

var badFillColorType = new Array();
var badFillColorIndex = 0;

// Keep track if we are accidentally in RGB color mode
var badColorMode = 0;

// Check document color space
if ( theDoc.documentColorSpace == DocumentColorSpace.CMYK ) {
	//all good
	debugln( "Document color space is CMYK." )
} else {
	debugln( "Document color space is RGB." )
	badColorMode = 1;
}

var dlg;


// Loop thru all text frames in document	
for ( var i = 0; i < theDoc.textFrames.length; i ++ ) {

	var item = theDoc.textFrames[i];

	debugln( "Starting processing of text frame # " + i );
	debugln( "  Name:         " + item.name );
	debugln( "  Type:         " + item.typename );
	debugln( "  Parent name:  " + item.parent.name );
	debugln( "  Parent type:  " + item.parent.typename );
	//debugln( "  Contents:     " + item.contents );
 	
	if ( item.typename == 'TextFrame' ) {
   	
   	    debugln( '' );

  		var badType = "";
   	
   		//for ( var ci = 0; ci < item.characters.length; ci++ ) {
   		//}
   	
   		var inBadBlackType = false;
		var inBadOverprintType = false;
		var inBadFillColorType = false;

   	
   		for ( var j = 0; j < item.textRanges.length; j++ ) {
   			
   			var tr = item.textRanges[j];
   			
   			// xdebugln( 'text range is ' + tr.contents );

            
            
            var fillColor;
            try {
                xdebugln( tr.contents + ' ' + inBadFillColorType + ' ' + badFillColorIndex);
			    fillColor = tr.characterAttributes.fillColor;
               	if ( inBadFillColorType ) {
               	    xdebugln ( '--------- i am here --------- ' );
                 	inBadFillColorType = false;
                 	badFillColorIndex++;
                 }
            } catch ( e ) {
             	if ( ! inBadFillColorType ) {
   					inBadFillColorType = true;
   					badFillColorType[ badFillColorIndex ] = new Object();
   					badFillColorType.contents = "";
   				} 
   				
   				if ( typeof( badFillColorType[badFillColorIndex].contents )=="undefined" ) {
   					badFillColorType[badFillColorIndex].contents = tr.contents;
   				} else {	
   					badFillColorType[badFillColorIndex].contents += tr.contents;   				
   				}
   				
                continue;
            } 

   			//var fillColor = tr.characterAttributes.fillColor;

   			
   			
   			var c = undefined;
   			var m = undefined;
   			var y = undefined;
   			var k = undefined;
   			
   			
   			// we have CMYK type.
   			if ( fillColor.typename == "CMYKColor" ) {
   				c = fillColor.cyan;		
   				m = fillColor.magenta;		
   				y = fillColor.yellow;		
   				k = fillColor.black;	
   					
   			} else if ( fillColor.typename == "SpotColor" ) {
   				
   				var theSpot = fillColor.spot;
   				
   				if ( theSpot.color.typename == "CMYKColor" ) {
	   				c = theSpot.color.cyan;		
 	  				m = theSpot.color.magenta;		
  	 				y = theSpot.color.yellow;		
   					k = theSpot.color.black;		
   				
     			}
   			}  else if ( fillColor.typename == "GrayColor" ) {
   					c = 0; m = 0; y = 0; 
   					k = fillColor.gray;
   			} 			
   			
   			// Check for bad type
   			if ( k > 60 && c+m+y > 0 ) {
   				if ( ! inBadBlackType ) {
   					inBadBlackType = true;
   					badBlackType[ badBlackIndex ] = new Object();
   					badBlackType.contents = "";
   				}
   				if ( typeof( badBlackType[badBlackIndex].contents )=="undefined" ) {
   					badBlackType[badBlackIndex].contents = tr.contents;
   				} else {	
   					badBlackType[badBlackIndex].contents += tr.contents;   				
   				}
				badBlackType[badBlackIndex].cmyk     = "" + c + " " + m + " " + y + " " + k;
				inBadBlackType = true;
				
   			} else {
   				if ( inBadBlackType == true ) {
   					inBadBlackType = false;
   					badBlackIndex++;
   				}
   			}
  
   			// Check for overprint white type. 			
   			if ( ( c + m + y + k ) == 0  && tr.characterAttributes.overprintFill ) {
   				if ( ! inBadOverprintType ) {
   					inBadOverprintType = true;
   					badOverprintType[ badOverprintIndex ] = new Object();
   					badOverprintType.contents = "";
   				}
   				if ( typeof( badOverprintType[badOverprintIndex].contents )=="undefined" ) {
   					badOverprintType[badOverprintIndex].contents = tr.contents;
   				} else {	
   					badOverprintType[badOverprintIndex].contents += tr.contents;   				
   				}
				badOverprintType[badOverprintIndex].cmyk     = "" + c + " " + m + " " + y + " " + k;
				inbadOverprintType = true;
				
   			} else {
   				if ( inBadOverprintType == true ) {
   					inBadOverprintType = false;
   					badOverprintIndex++;
   				}   			
   			}
   			
   			
   			
   			debugln( "--> " + tr.contents + " is " + fillColor.typename + " " + c + " " + m + " " + y + " " + k );
   			
   			
   			
   		
   		}
   		
   		// out of textframe. reset inFlag. increment index by 1 if still in;
   		if ( inBadBlackType ) {
   			badBlackIndex++;
   		}

   		// out of textframe. reset inFlag. increment index by 1 if still in;
   		if ( inBadOverprintType ) {
   			badOverprintIndex++;
   		}
   		
   		if ( inBadFillColorType ) {
   			badFillColorIndex++;
   		}
   		
   		


   	
     			
   	  					//item.contents = mapData[dataType][ key ].text; 				
						//item.textRange.characterAttributes.fillColor = getColorfromString( mapData[dataType][ key ].fillCMYK );
     	
   	}
}

var numProblems = badColorMode + badBlackType.length + badOverprintType.length + badFillColorType.length;

var resultsMsg = "";

if ( numProblems > 0 ) {


	resultsMsg +=  "Summary of problems found in \"" + theDoc.name + "\"\n";
	resultsMsg +=  '\n';

	if ( badColorMode ) {
	
		alertMsg += "\nYour document is in RGB color mode. Change it to CMYK.\n"
	
		resultsMsg +=  'BAD DOCUMENT COLOR MODE\n';
		resultsMsg +=  '====================================================================\n';
		resultsMsg +=  "Your document is RGB color mode. You must change it to CMYK.\n";
		resultsMsg +=  '\n';
	}

	if ( badBlackType.length  ) {
	
	
		alertMsg += "\nYour document has " +  badBlackType.length + " pieces of 4-color black type." 
		
		resultsMsg +=  'POSSIBLE 4-COLOR BLACK TYPE\n';
		resultsMsg +=  '================================================\n';

		for ( var i=0; i<badBlackType.length; i++ ) {
			debugln("");
			debugln("-------------------------");
			debugln("Bad 4-color black type")
			debugln( badBlackType[i].contents );
			debugln( badBlackType[i].cmyk );
			
			resultsMsg += badBlackType[i].contents + "\n";
			
		}
		resultsMsg +=  '\n';
	}
	
	
	
	
	
	if ( badFillColorType.length  ) {
	
		alertMsg += "\nYour document has " +  badFillColorType.length + " pieces of type with invalid CMYK specs." 
		
		
		resultsMsg +=  'POSSIBLE TYPE WITH INVALID CMYK SPECS\n';
		resultsMsg +=  '(It might be OK, but check each piece by hand, please)\n';
		resultsMsg +=  '================================================\n';

		for ( var i=0; i<badFillColorType.length; i++ ) {
			debugln("");
			debugln("-------------------------");
			debugln(" type")
			debugln( badFillColorType[i].contents );
			debugln( badFillColorType[i].cmyk );
			
			resultsMsg += badFillColorType[i].contents + "\n";
			
		}
		resultsMsg +=  '\n';
	}

	if ( badOverprintType.length  ) {
	
		alertMsg += "\nYour document has " +  badOverprintType.length + " pieces of overprinting white type." 

		resultsMsg +=  'POSSIBLE OVERPRINTING WHITE TYPE\n';
		resultsMsg +=  '================================================\n';

		for ( var i=0; i<badOverprintType.length; i++ ) {

			debugln("");
			debugln("-------------------------");
			debugln("Overprinting white type")
			debugln( badOverprintType[i].contents );

			resultsMsg += badOverprintType[i].contents + "\n";
		}
		resultsMsg += '\n';
	}
	
	

} else {
	resultsMsg = 'Your document appears to be OK.';
}

displayAlert( resultsMsg );


function displayAlert( msg ) {
	dlg = new Window('dialog', 'Preflight Report'); 
	
	var msgPnl = dlg.add('panel', undefined, 'Preflight Report'); 
	msgPnl.orientation = 'row';

	msgPnl.report = msgPnl.add('edittext', undefined, msg,{multiline:true});
	msgPnl.report.preferredSize = [600,400];




	var btnPnl = dlg.add('group', undefined, ''); 
	btnPnl.orientation = 'row'
	btnPnl.cancelBtn = btnPnl.add('button', undefined, 'Cancel', {name:'cancel'});
	btnPnl.cancelBtn.onClick = function() { dlg.close() };

	dlg.show();
}
 

app.redraw();


function getCMYKfromString( s ) {
     		var cmyk      = s.split( "," );
     		
     		var color     = new CMYKColor();
     		color.cyan    = cmyk[0];
     		color.magenta = cmyk[1];
			color.yellow  = cmyk[2];
			color.black   = cmyk[3];
			
			return color;
}	


function getColorfromString( s ) {

			var colorType = '';
			var colorData = '';
			
			if ( s.match( ":" ) ) {
				var colorParts = s.split( ':' );
				colorType = colorParts[0];
				colorData = colorParts[1];
			} else {
				colorType = 'CMYK';
				colorData = s;
			}

			var color; 
		
			debugln( '    --> Color type: '  + colorType ) ;
			debugln( '    --> Color data: '  + colorData ) ;

					
			if ( colorType == 'CMYK' ) {
    	 		var cmyk      = s.split( "," );
     		
     			color     = new CMYKColor();
     			color.cyan    = cmyk[0];
     			color.magenta = cmyk[1];
				color.yellow  = cmyk[2];
				color.black   = cmyk[3];
				
			} else if ( colorType == 'Swatch' ) {
				try {
					if ( theDoc.swatches[colorData] ) {
						var swatch = theDoc.swatches[colorData];
						color = swatch.color;
					}
				} catch (e) {
					alert( "The swatch " + colorData + " doesn’t exist" );
				}
			}
			
			return color;
}	

true;

//////////////////////////////////////////
// Functions start here

function getNewName(sourceDoc, destFolder) {
	var docName = sourceDoc.name;
	var ext = '.jpg'; // new extension for pdf file
	var newName = "";

	// if name has no dot (and hence no extension,
	// just append the extension
	if (docName.indexOf('.') < 0) {
		newName = docName + ext;
	} else {
		var dot = docName.lastIndexOf('.');
		newName += docName.substring(0, dot);
		newName += ext;
	}

	newName = newName.replace( / /g, '-' );
	
	
	// Create a file object to save the pdf
	saveInFile = new File( destFolder + '/' + newName );
	return saveInFile;
}


function debugln ( s ) {
	if ( debugThisScript ) {
		$.writeln( s );
	}
}

function xdebugln ( s ) {
	if ( xdebugThisScript ) {
		$.writeln( s );
	}
}