 { 
   split($4,date,"/");
   printf "%s\t%d/%02d/%02d\t%s\n",$1, date[3], date[1], date[2], $3
 } 