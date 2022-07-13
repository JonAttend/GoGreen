// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router')();
const graph = require('../graph.js');
const addDays = require('date-fns/addDays');
const formatISO = require('date-fns/formatISO');
const startOfWeek = require('date-fns/startOfWeek');
const zonedTimeToUtc = require('date-fns-tz/zonedTimeToUtc');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');


function FileConvertSize(aSize){
	aSize = Math.abs(parseInt(aSize, 10));
	var def = [[1, 'octets'], [1024, 'ko'], [1024*1024, 'Mo'], [1024*1024*1024, 'Go'], [1024*1024*1024*1024, 'To']];
	for(var i=0; i<def.length; i++){
		if(aSize<def[i][0]) return (aSize/def[i-1][0]).toFixed(2)+' '+def[i-1][1];
	}
}

function FileNameExtention(name){
  const regex = /(?:\.([^.]+))?$/;
  return regex.exec(name)[1];
}

// <GetRouteSnippet>
/* GET /calendar */
router.get('/',
  async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { drive: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      
      try {
        // Get the events
        const myDrive = await graph.getDriveView(
          req.app.locals.msalClient,
          req.session.userId);

        console.log("mydrive ", myDrive);
        myDrive.value.map( item => {
          
          if(item.file) {
            
            const ext = FileNameExtention(item.name);
            
            switch (ext) {
              case 'docx':
                item.word = ext;
                break;
              case 'pptx':
                item.powerpoint = ext;
                break;
              case 'xlsx':
                item.excel = ext;
                break;
              case 'vsdx':
                item.drawing = ext;
                break;
              case 'pdf':
                item.pdf = ext;
                break;
              case 'PNG':
              case 'jpg':
              case 'jpeg':
                item.image = ext;
                break;
              default:
                console.log(`Sorry, we do not know this extension : ${ext}.`);
            }
          } else if (item.package && item.package.type === 'oneNote') {
            item.onenote = item.package.type;
          }
          if(item.size) {
            item.size = FileConvertSize(item.size);
          }
        })
        // Assign data drive to the view parameters
        params.drive = myDrive.value;
        // console.log(params.drive)
      } catch (err) {
        console.error(err)
        req.flash('error_msg', {
          message: 'Could not fetch events',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }

      res.render('drive', params);
    }
  }
);
// </GetRouteSnippet>

module.exports = router;
