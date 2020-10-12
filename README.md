# Quicker

    Employee.myGreet = function(msg, cb) {
        cb(null, 'Greetings... ' + msg);
    }

    Employee.greet = async function(msg) {
        return 'Greetings... ' + msg;
    }
	
    Employee.greet = async (msg) => 'Greetings... ' + msg;
