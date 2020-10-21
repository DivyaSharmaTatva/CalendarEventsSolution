'use strict';

const build = require('@microsoft/sp-build-web');
require('./gulpfile-spcaf');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(require('gulp'));
