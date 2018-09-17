var Connection = require('tedious').Connection;
var Request = require('tedious').Request;
const  http= require('http');
const sql = require('mssql');
const hostname = '127.0.0.1';
const port = 5000;

const server = http.createServer((req, res) => {
    res.statusCode = 200;
    ///res.setHeader('Content-Type', 'text/plain');
   // res.end('Hello World\n');
});

server.listen(port, hostname, () => {
    console.log(`Server running at http://${hostname}:${port}/`);
});

// Create connection to database
var config =
    {
        userName: 'guy', // update me
        password: 'Wilks2003', // update me
        server: 'winedb.database.windows.net', // update me
        port: '1433',
        options:
            {
                database: 'winedb' //update me
                , encrypt: true
            }
    }
var connection = new Connection(config);

// Attempt to connect and execute queries if connection goes through
connection.on('connect', function(err)
    {
        if (err)
        {
            console.log(err)
        }
        else
        {
            console.log('connected');
            var request = new sql.Request();

            request.query(
                "SELECT * FROM wineTest",
                function(err, rows)
                {
                    console.log(rows);
                }
            );
        }
    }
);

