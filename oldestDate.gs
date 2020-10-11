function oldestDate(date1,date2,date3) {
  var oldest = date1;
  
  if (date1 > date2){
    oldest = date2;
  }
  
  if (date2 > date3){
    oldest = date3;
  }
  
  return oldest;
  
}
