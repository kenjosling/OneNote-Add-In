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
                    .set("Authorization", "Bearer INSERT_TOKEN_HERE")
                    .then(function(response) {
                        var page = context.application.getActivePage();
                        var outline = page.addOutline(240, 240, "It worked!");
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
