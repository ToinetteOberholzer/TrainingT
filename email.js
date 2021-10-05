const Imap = require('imap'), inspect = require('util').inspect;
var fs      = require('fs');
var base64  = require('base64-stream');
const opDir = 'C:/Users/Sintrex Training/Desktop/PJ2/Attachment/';


var imap    = new Imap({
  directory: opDir,
  user: 'toinette.oberholzer@sintrex.com',
  password: 'Roxanne81!',
  host: 'outlook.office365.com',
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false }
  //,debug: function(msg){console.log('imap:', msg);}
});

function toUpper(thing) { return thing && thing.toUpperCase ? thing.toUpperCase() : thing;}

function findAttachmentParts(struct, attachments) {
  attachments = attachments ||  [];
  for (var i = 0, len = struct.length, r; i < len; ++i) {
    if (Array.isArray(struct[i])) {
      findAttachmentParts(struct[i], attachments);
    } else {
      if (struct[i].disposition && ['INLINE', 'ATTACHMENT'].indexOf(toUpper(struct[i].disposition.type)) > -1) {
        attachments.push(struct[i]);
      }
    }
  }
  return attachments;
}

function buildAttMessageFunction(attachment) {
  var filename = attachment.params.name;
  var encoding = attachment.encoding;

  return function (msg, seqno) {
    var prefix = '(#' + seqno + ') ';
    /*msg.on('body', function(stream, info) {
      //Create a write stream so that we can stream the attachment to file;
      console.log(prefix + 'Streaming this attachment to file', filename, info);
      var writeStream = fs.createWriteStream(filename);
      writeStream.on('finish', function() {
        console.log('Done writing to file %s', filename);
      });

      //stream.pipe(writeStream); this would write base64 data to the file.
      //so we decode during streaming using 
      if (toUpper(encoding) === 'BASE64') {
        //the stream is base64 encoded, so here the stream is decode on the fly and piped to the write stream (file)
        //stream.pipe(base64.decode()).pipe(writeStream);
        stream.pipe(writeStream);
      } else  {
        //here we have none or some other decoding streamed directly to the file which renders it useless probably
        //stream.pipe(writeStream);
        console.log('Error reading');
      }
    });
    msg.once('end', function() {
      console.log('Finished attachment %s', filename);
    });*/
  };
}



function openInbox(cb) {
    imap.openBox('INBOX', true, cb);
  }
   
  imap.once('ready', function() {
    openInbox(function(err, box) {
      if (err) throw err;
      var f = imap.seq.fetch('1:3', {
      //  bodies: 'HEADER.FIELDS (FROM TO SUBJECT DATE)',
      bodies: 'HEADER.FIELDS (FROM SUBJECT)',
        struct: true
      });
      f.on('message', function(msg, seqno) {
      // console.log('Message #%d', seqno);

     // var prefix = '(#' + seqno + ') ';
      msg.on('body', function(stream, info) {
        var buffer = '';
        stream.on('data', function(chunk) {
          buffer += chunk.toString('utf8');
        });
        stream.once('end', function() {
          console.log('\n------------------------------------------------')
          console.log('                   Email Received:                ')
          //console.log('Parsed header: %s', inspect(Imap.parseHeader(buffer)));
          console.log(seqno, inspect(Imap.parseHeader(buffer)));
          
        });
      });
     
      msg.once('attributes', function(attrs) {
        var attachments = findAttachmentParts(attrs.struct);
       // console.log('Email: #'+ seqno + ' has attachments: %d', attachments.length);
        for (var i = 0, len=attachments.length ; i < len; ++i) {
          var attachment = attachments[i];
          
          console.log('Attachments:', attachment.params.name);
          var f = imap.fetch(attrs.uid , { //do not use imap.seq.fetch here
            bodies: [attachment.partID],
            struct: true
          });
          //build function to process attachment message
          f.on('message', buildAttMessageFunction(attachment));
         // var emm = [{from: '.com', subject: '', attachSave: attachment}];
        }
      });
      msg.once('end', function() {
        console.log('\nFinished email');
      });
    });
    f.once('error', function(err) {
      console.log('Fetch error: ' + err);
    });
    f.once('end', function() {
      console.log('Done fetching all messages!');
      imap.end();
    });
  });
});

imap.once('error', function(err) {
  console.log(err);
});

imap.once('end', function() {
  console.log('Connection ended');
});

imap.connect();



//module.exports = email;