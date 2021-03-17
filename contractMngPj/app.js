var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');


//�����큫
var update = require('./routes/update');
//�ق��聫
var contractDel = require('./routes/contractDel');
var importpro = require('./routes/Import_Process');





var search_list = require('./routes/search_manager');
var writeExcel = require('./routes/writeExcel');

var app = express();

var bodyParser = require('body-parser');

app.use(bodyParser.urlencoded({
  extended: false,
  parameterLimit: 10000,
  limit: 1024 * 1024 * 10
}));
app.use(bodyParser.json({
  extended: false,
  parameterLimit: 10000,
  limit: 1024 * 1024 * 10
}));



// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', indexRouter);
app.use('/users', usersRouter);

app.use('/search_list', search_list);
app.use('/writeExcel', writeExcel);

//�����큫
app.use('/update', update);
//�ق��聫
app.use('/delete', contractDel);
app.use('/import', importpro);


// catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
