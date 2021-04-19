/*
'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(require('gulp'));
*/
'use strict';

// check if gulp dist was called
if (process.argv.indexOf('dist') !== -1) {
  // add ship options to command call
  process.argv.push('--ship');
}

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const gulpSequence = require('gulp-sequence');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Create clean distrubution package
gulp.task('dist', gulpSequence('clean', 'bundle', 'package-solution'));
// Create clean development package
gulp.task('dev', gulpSequence('clean', 'bundle', 'package-solution'));

/********************************************************************************************
 * Adds an alias for handlebars in order to avoid errors while gulping the project
 * https://github.com/wycats/handlebars.js/issues/1174
 * Adds a loader and a node setting for webpacking the handlebars-helpers correctly
 * https://github.com/helpers/handlebars-helpers/issues/263
 ********************************************************************************************/
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {

    generatedConfiguration.resolve.alias = { handlebars: 'handlebars/dist/handlebars.min.js' };
/*
    generatedConfiguration.module.rules.push(
      { test: /\.js$/, loader: 'unlazy-loader' }
    );
*/
    generatedConfiguration.node = {
      fs: 'empty'
    }

    return generatedConfiguration;
  }
});
build.initialize(gulp);
