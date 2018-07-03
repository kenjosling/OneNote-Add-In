/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as OfficeHelpers from '@microsoft/office-js-helpers';
import request from 'superagent';

$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
    try {
            await OneNote.run(async context => {
                /**
                 * Insert your OneNote code here
                 */
                request
                    .get('https://graph.microsoft.com/beta/me/photos/240x240/$value')
                    .set("Authorization", "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEWDhHQ2k2SnM2U0s4MlRzRDJQYjdydlAxbllBclN6YS1OT1c5V3pyV19NZHkzampKVHFPR1h2WDd5ZFFkWnRISzJMM1FSeUdQbFNGRER5N3pKWUxVMzkybkU2VXExMWF2TWNIQkpuZHUwLUNBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIiwia2lkIjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yM2FlNTdkZi04YTUxLTQ0NGEtYTU5YS0zMjgxMTg3MDVlZmMvIiwiaWF0IjoxNTMwNTc2OTY5LCJuYmYiOjE1MzA1NzY5NjksImV4cCI6MTUzMDU4MDg2OSwiYWNyIjoiMSIsImFpbyI6IkFVUUF1LzhIQUFBQTcwWmFuUFcxNHlRaGxFbFdyaEV0U0p0L0EvZW1laFpaNW1LcEF6Z2FiRWYwZ3RjdTVMTldTdFZuOWhYTE5mOWVRMXNQcGNTYlV4Vm81SnlVWDU5M0tRPT0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIGV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkFsYmFueSIsImdpdmVuX25hbWUiOiJCcnVjZSIsImlwYWRkciI6IjE0NC4xMzAuMTczLjE2MiIsIm5hbWUiOiJCcnVjZSBBbGJhbnkiLCJvaWQiOiI2MTBhNzM5YS0yNzE5LTQ1ODQtYjNkNy00YzkxYTg0NGU5ZjYiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMzI3MDI4MDIyMy01MjAyMjQ5MDEtNDYwNzYyMTc4LTE1MTA2IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDM3RkZFOUJDMTdCOTIiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBGaWxlcy5SZWFkV3JpdGUuQWxsIE1haWwuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgUGVvcGxlLlJlYWQgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUYXNrcy5SZWFkV3JpdGUgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIiwic3ViIjoiTFNTRTBpdGxQZXNDcjdaZUpJMDBlMUZVWlRybVRQbFFRWlY4SXJiSTFIYyIsInRpZCI6IjIzYWU1N2RmLThhNTEtNDQ0YS1hNTlhLTMyODExODcwNWVmYyIsInVuaXF1ZV9uYW1lIjoiYnJ1Y2UuYWxiYW55QGtsb3VkLmNvbS5hdSIsInVwbiI6ImJydWNlLmFsYmFueUBrbG91ZC5jb20uYXUiLCJ1dGkiOiJhWjF6WUh4WUpFT2RCMEFiMVkweEFBIiwidmVyIjoiMS4wIn0.TMr-Px09bf_4JuNhtXFBgmZc5P4-iymkdmCBCc4WHrdhDF9aIV4bdWOyZxVFV8D9au_qwPlDRxZTrp0ekvbWHrKySR6ixbFMKIzYPy5f25WNSUwlk-MtICVr6zynG2_uf9X7XNdP2B5V4hTNnO_nuwed1M-KGhLjwCbEuaWoAGRiDmCK_9cYiBcQ3UiD1JD_0nmG2OAllaCdyF2MoH9ZShXboqPPTrMuTCFR4Pqkr7kDqFibQKphFAB0BAg3gCXRbozVzRcURNMlGhaCrZI4geHzLIYqu4c5_KQqqeQGgtwxcw97dbxIvhUlBRJaDUL1PuUMd2sH_N12NmdBB5cYfw")
                    .responseType('arraybuffer')
                    .then(function(response) {
                        var page = context.application.getActivePage();
                        var outline = page.addOutline(240, 240, "It worked!");
                        var base64EncodedStr = Buffer.from(response.body).toString('base64');
                        outline.appendImage(base64EncodedStr, 240, 240);
                        return context.sync();
                    }, function(error) {
                        var page = context.application.getActivePage();
                        var outline = page.addOutline(240, 240, "It didn't work");
                        return context.sync();
                    });
            });
        } catch(error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        };
}
