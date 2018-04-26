function myFunction() {
  var regex = new RegExp('^O[0-9]{5}.*[0-9]$');
  var str = "TUE";
  Logger.log(regex.test(str))
}
