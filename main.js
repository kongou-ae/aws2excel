var lambda = require('./index.js');

var event = {
  "operation": "echo",
  "message": "Hello world!"
};

var context = {};
context.done = function(a, b) {
	if(a){console.log("ERROR:"+a);
	}else{console.log("SUCCESS");}
	console.log(b);
};

lambda.handler(event, context);