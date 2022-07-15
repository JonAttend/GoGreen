// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router')();
const { formatDistance, parseISO  } = require('date-fns');
const graph = require('../graph.js');

function ConvertFileSize(fileSizeInBytes) {
  var i = -1;
  var byteUnits = [' kB', ' MB', ' GB', ' TB', 'PB', 'EB', 'ZB', 'YB'];
  do {
      fileSizeInBytes = fileSizeInBytes / 1024;
      i++;
  } while (fileSizeInBytes > 1024);

  return Math.max(fileSizeInBytes, 0.1).toFixed(1) + byteUnits[i];
};

function ConvertLastTimeModify(date){
  return formatDistance(new Date(), parseISO(date), { includeSeconds: true }) 
}

function FileNameExtention(name){
  const regex = /(?:\.([^.]+))?$/;
  return regex.exec(name)[1];
}

function CheckFileType(obj){
  if(obj.name) {
    const ext = FileNameExtention(obj.name);
    switch (ext) {
      case 'docx':
        obj.word = ext;
        break;
      case 'pptx':
        obj.powerpoint = ext;
        break;
      case 'xlsx':
        obj.excel = ext;
        break;
      case 'vsdx':
        obj.drawing = ext;
        break;
      case 'pdf':
        obj.pdf = ext;
        break;
      case 'PNG':
      case 'jpg':
      case 'jpeg':
        obj.imageType = ext;
        break;
      default:
        console.log(`Sorry, we do not know this extension : ${ext}.`);
    }
    return obj;
  }
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
        // Get the drive
        const myDrive = await graph.getDriveView(
          req.app.locals.msalClient,
          req.session.userId);

          console.log("mydrive ", myDrive);
        
          myDrive[1].value.map( item => {
          
          if(item.file) {
            // Add property 'type file' in current obj for hbs view
            CheckFileType(item);
          } else if (item.package && item.package.type === 'oneNote') {
            item.onenote = item.package.type;
          }

          if(item.size) {
            item.size = ConvertFileSize(item.size);
          }

          if(item.lastModifiedDateTime) {
            // take the last date modify and return "about 3 hours"
            item.lastModifiedDateTime = ConvertLastTimeModify(item.lastModifiedDateTime);
          }

        })
        // Convert octet size for remaining / total / used
        for (const property in myDrive[0].quota) {
          if(property === 'remaining' || property === 'total' || property === 'used' ){
            myDrive[0].quota[property] = ConvertFileSize(myDrive[0].quota[property]);
          }
        }

        // Assign data drive to the view parameters
        params.dashboard = myDrive[0].quota;
        params.drive = myDrive[1].value;
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
