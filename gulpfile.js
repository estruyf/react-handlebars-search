'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const path = require('path');

/**
 * Task is necessary for VSTS
 */
const envCheck = build.subTask('environmentCheck', (gulp, config, done) => {
    /********************************************************************************************
    * Adds an alias for handlebars in order to avoid errors while gulping the project
    * https://github.com/wycats/handlebars.js/issues/1174
    * Adds a loader and a node setting for webpacking the handlebars-helpers correctly
    * https://github.com/helpers/handlebars-helpers/issues/263
    ********************************************************************************************/
    build.configureWebpack.mergeConfig({
      additionalConfiguration: (generatedConfiguration) => {
        generatedConfiguration.resolve.alias = {
          handlebars: 'handlebars/dist/handlebars.min.js'
        };

        // Check if running in debug or ship
        if (config.production) {
          generatedConfiguration.module.rules.push({
            test: /\.js$/,
            loader: 'unlazy-loader'
          });
        } else {
          generatedConfiguration.module.rules.push({
            test: /\.js$/,
            loader: 'unlazy-loader',
            exclude: [path.resolve(__dirname, './lib/')]
          });
        }


        generatedConfiguration.node = {
          fs: 'empty'
        }

        return generatedConfiguration;
      }
    });

    done();
});
build.rig.addPreBuildTask(envCheck);

build.initialize(gulp);
