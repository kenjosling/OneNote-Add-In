var request = require("request");
var fs = require('fs');

request.get({
    "headers": { 
        "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEWDhHQ2k2SnM2U0s4MlRzRDJQYjdyU2YxVTlMMXRBdnRmQjlzUTJwSUtSNkpVWXBpNFdHcDY3U3g0eVV0MGpiSks2LWZjN09tRGo2RUVpSk15R1JGS29VdElsV2tUbTBmYThhQndkLUZPakNBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIiwia2lkIjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yM2FlNTdkZi04YTUxLTQ0NGEtYTU5YS0zMjgxMTg3MDVlZmMvIiwiaWF0IjoxNTMwNDkyMzgwLCJuYmYiOjE1MzA0OTIzODAsImV4cCI6MTUzMDQ5NjI4MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhIQUFBQTBHNW1vbkQvWXVZaDFmTmMrZktmNTM2d3FkSkFCQjI1U3NRRnRtcXgybmFzOUZyOXFpck5uQkhQdUM2Q2hnMmx0elI1SXVKWnpobit0UzdDWUtXVjlLSjVpZUV3bVZmSVFLc05MejNROU13PSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiQmx1ZSBTdGVlbCIsImFwcGlkIjoiNjY4OWY5ZjYtNDMxYy00ODc5LWIwNDktOWU4YzQzZjhhZWM3IiwiYXBwaWRhY3IiOiIxIiwiZGV2aWNlaWQiOiI5NzFhODg4ZS1jM2I2LTRmN2EtYjQxMi0xNGI2Yzg3YjA3NDIiLCJmYW1pbHlfbmFtZSI6Ikpvc2xpbmciLCJnaXZlbl9uYW1lIjoiS2VuIiwiaXBhZGRyIjoiMS4xMjkuMTA4LjU4IiwibmFtZSI6IktlbiBKb3NsaW5nIiwib2lkIjoiNGNiNjg5OTAtMWE2My00YmY2LTg5YjAtMWE3MmIwNGM4ZWU2Iiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTMyNzAyODAyMjMtNTIwMjI0OTAxLTQ2MDc2MjE3OC0xNDk1MiIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMDAwMDk2M0NFNEJEIiwic2NwIjoiVXNlci5SZWFkIiwic2lnbmluX3N0YXRlIjpbImttc2kiXSwic3ViIjoiYldKZjJfb1c2RmVQLUZGMnJSTHZSYWZlQmVMODZvb0pidG1lVEhwd1JISSIsInRpZCI6IjIzYWU1N2RmLThhNTEtNDQ0YS1hNTlhLTMyODExODcwNWVmYyIsInVuaXF1ZV9uYW1lIjoia2VuLmpvc2xpbmdAa2xvdWQuY29tLmF1IiwidXBuIjoia2VuLmpvc2xpbmdAa2xvdWQuY29tLmF1IiwidXRpIjoiRDdsTFpmWk5Ga2F2SGhwaUlIWW5BQSIsInZlciI6IjEuMCJ9.N4HOHyG5LbR8zoqsJsixN5pKLW9EKWQBZByb40f6SqgsljRQ6ICzAs5--Wk0FKxfgKohBPilCf0-_oJRAb0yedDa_2GKJvAFLxyiudJMv3BOcZ0aI4dXxggpTNI-udhY1zW5-0us7fY8z6pQLHx4gAJJVlMk_jVzvkxA2t9d94Itkq5vupbAB6gq3Ryoqvs0Hkui6q4MRCVncM-1SLlOmfXMSmXVD3ZGdWkTWNueFO_IQtOD28IbMH0ZsJHERpsLZ1kC4F0fcOeyxKH-OyAC0xmL0h99f3wsdUp7yX4GY-V9GrLZh2KbO6o2Gs0qjIwT281MJhQyL8qB-4fktqSdxg" },
        "url": "https://graph.microsoft.com/beta/me/photos/240x240/$value"
}   , (error, response, body) => {
        if(error) {
            return console.dir(error);
        }
        fs.writeFile("./mug.txt", body, function(err) {
            if(err) {
                return console.log(err);
            }
        }); 
})
