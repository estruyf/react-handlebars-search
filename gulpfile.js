'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

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
            test: [/(\/handlebars\/)([A-Za-z0-9\-\.\/]+)(\.js$)/,
            /(\/handlebars-helpers\/)([A-Za-z0-9\-\.\/]+)(\.js$)/,
            /(\/create-frame\/)([A-Za-z0-9\-\.\/]+)(\.js$)/],
            loader: 'unlazy-loader'
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
