'use strict';

// Variables
var async = require('async');
var excelbuilder = require('msexcel-builder');
var request = require('request-json');
var fs = require('fs');
var urlHitTask = [];
var xlResultList = [];

// Constants
const apiKey = 'AIzaSyDkEX-f1JNLQLC164SZaobALqFv4PHV-kA';
const strategyType = 'mobile';
const googlePageSpeedUrl = 'https://www.googleapis.com/pagespeedonline/';

var siteUrls = fs.readFileSync('./urlList.txt', 'utf8');
siteUrls = siteUrls.split('\n');

siteUrls.forEach((siteUrl)=> {
    urlHitTask.push((callback)=> {
        (function (siteUrl) {
            var client = request.createClient(googlePageSpeedUrl);
            client.get('v3beta1/runPagespeed?url=' + siteUrl + '&strategy=' + strategyType + '&key=' + apiKey,
                function (err, res, body) {
                    /*console.log(siteUrl);
                     console.log(body.formattedResults.ruleResults.AvoidInterstitials.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.AvoidLandingPageRedirects.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.AvoidPlugins.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.ConfigureViewport.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.EnableGzipCompression.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.LeverageBrowserCaching.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.MainResourceServerResponseTime.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.MinifyCss.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.MinifyHTML.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.MinifyJavaScript.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.MinimizeRenderBlockingResources.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.OptimizeImages.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.PrioritizeVisibleContent.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.SizeContentToViewport.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.SizeTapTargetsAppropriately.localizedRuleName);
                     console.log(body.formattedResults.ruleResults.UseLegibleFontSizes.localizedRuleName);
                     */
                    xlResultList.push({
                        siteUrl: siteUrl,
                        AvoidInterstitials: body.formattedResults.ruleResults.AvoidInterstitials.localizedRuleName,
                        AvoidLandingPageRedirects: body.formattedResults.ruleResults.AvoidLandingPageRedirects.localizedRuleName,
                        AvoidPlugins: body.formattedResults.ruleResults.AvoidPlugins.localizedRuleName,
                        ConfigureViewport: body.formattedResults.ruleResults.ConfigureViewport.localizedRuleName,
                        EnableGzipCompression: body.formattedResults.ruleResults.EnableGzipCompression.localizedRuleName,
                        LeverageBrowserCaching: body.formattedResults.ruleResults.LeverageBrowserCaching.localizedRuleName,
                        MainResourceServerResponseTime: body.formattedResults.ruleResults.MainResourceServerResponseTime.localizedRuleName,
                        MinifyCss: body.formattedResults.ruleResults.MinifyCss.localizedRuleName,
                        MinifyHTML: body.formattedResults.ruleResults.MinifyHTML.localizedRuleName,
                        MinifyJavaScript: body.formattedResults.ruleResults.MinifyJavaScript.localizedRuleName,
                        MinimizeRenderBlockingResources: body.formattedResults.ruleResults.MinimizeRenderBlockingResources.localizedRuleName,
                        OptimizeImages: body.formattedResults.ruleResults.OptimizeImages.localizedRuleName,
                        PrioritizeVisibleContent: body.formattedResults.ruleResults.PrioritizeVisibleContent.localizedRuleName,
                        SizeContentToViewport: body.formattedResults.ruleResults.SizeContentToViewport.localizedRuleName,
                        SizeTapTargetsAppropriately: body.formattedResults.ruleResults.SizeTapTargetsAppropriately.localizedRuleName,
                        UseLegibleFontSizes: body.formattedResults.ruleResults.UseLegibleFontSizes.localizedRuleName
                    });
                    console.log('Site: '+siteUrl);

                    return callback(err, siteUrl);

                });
        })(siteUrl);
    });
});


async.series(urlHitTask, (err, result)=> {
    if (err) {
        throw err;
    } else {
        // Create a new workbook file in current working-path
        var workbook = excelbuilder.createWorkbook('./', 'pageSpeedResult.xlsx');

// Create a new worksheet with 10 columns and 12 rows
        var sheet1 = workbook.createSheet('sheet1', 17, xlResultList.length + 1);

        // Set Heading
        sheet1.set(1, 1, 'Site Url');
        sheet1.set(2, 1, 'Avoid Interstitials');
        sheet1.set(3, 1, 'Avoid Landing Page Redirects' );
        sheet1.set(4, 1, 'Avoid Plugins');
        sheet1.set(5, 1, 'Configure Viewport' );
        sheet1.set(6,  1, 'Enable Gzip Compression');
        sheet1.set(7,  1, 'Leverage Browser Caching');
        sheet1.set(8,  1, 'Main Resource Server Response Time');
        sheet1.set(9,  1, 'Minify Css');
        sheet1.set(10,  1, 'Minify HTML');
        sheet1.set(11,  1, 'Minify JavaScript');
        sheet1.set(12,  1, 'Minimize Render Blocking Resources');
        sheet1.set(13,  1, 'Optimize Images');
        sheet1.set(14,  1, 'Prioritize Visible Content');
        sheet1.set(15,  1, 'Size Content To Viewport');
        sheet1.set(16,  1, 'Size Tap Targets Appropriately');
        sheet1.set(17,  1, 'Use Legible FontSizes');

        // Fill some data
        for (var x = 0; x < xlResultList.length; x++) {
            sheet1.set(1, x + 2, xlResultList[x].siteUrl);
            sheet1.set(2, x + 2, xlResultList[x].AvoidInterstitials);
            sheet1.set(3, x + 2, xlResultList[x].AvoidLandingPageRedirects);
            sheet1.set(4, x + 2, xlResultList[x].AvoidPlugins);
            sheet1.set(5, x + 2, xlResultList[x].ConfigureViewport);
            sheet1.set(6, x + 2, xlResultList[x].EnableGzipCompression);
            sheet1.set(7, x + 2, xlResultList[x].LeverageBrowserCaching);
            sheet1.set(8, x + 2, xlResultList[x].MainResourceServerResponseTime);
            sheet1.set(9, x + 2, xlResultList[x].MinifyCss);
            sheet1.set(10, x + 2, xlResultList[x].MinifyHTML);
            sheet1.set(11, x + 2, xlResultList[x].MinifyJavaScript);
            sheet1.set(12, x + 2, xlResultList[x].MinimizeRenderBlockingResources);
            sheet1.set(13, x + 2, xlResultList[x].OptimizeImages);
            sheet1.set(14, x + 2, xlResultList[x].PrioritizeVisibleContent);
            sheet1.set(15, x + 2, xlResultList[x].SizeContentToViewport);
            sheet1.set(16, x + 2, xlResultList[x].SizeTapTargetsAppropriately);
            sheet1.set(17, x + 2, xlResultList[x].UseLegibleFontSizes);
        }

        // Save it
        workbook.save(function (err) {
            if (err)
                throw err;
            else
                console.log('congratulations, your Page Speed XL created');
        });

    }
});


