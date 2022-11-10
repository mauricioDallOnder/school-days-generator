/*

1)Get Holidays:
  a) Function that looks for holidays in a range of values ​​and places the days in an array of values

2)GetHolidays:
  b)Saturday job search function:

3)Get Calendar
  
  a) Receives a start and end date as a parameter
  b) creates an empty array that will receive the dates and return at the end of the function
  c) a function is created that adds days, and its parameter is the day value
  d) Create a constant that extracts the numeric value of the date object
  e) With this number you can use the setdate() function, as it will return the specific day of a given month,
  and make an increment with the days of the variable
  f) As long as the current date is less than the final date, the date array is incremented.
  g) returns the date array at the end of the function.
  ----------------------------------------------------
4) Function Calculate_dates:
  a) creates a function that calculates the school days on top of the generated calendar, this function receives as
  "Saturday" parameter, as Saturdays are different in the quarters.
  b) A constant called dates is created, which receives the function that generates the calendar
  (the function that generates the calendar receives 2 parameters, the start and end date).
  c) A "for" is done based on the generated calendar and the dates are placed inside an array.
  d) as the date is in GMT, it is necessary to extract only the part that has the day of the week and the date.
  e) Create an array, where all Saturdays and Sundays have been removed.
  However, if Saturday matches the Saturdays of previously defined arrays, it returns this value.
  f) Using the previously filtered Saturday array, another array is created that uses filter to filter the results
  --------------------------------------------------------------------------------------------------
5)Function searches for school days:
  a) If the day of the week receives "x" and the arr_datas in position i, (cut by the first 3 characters (in this case it is example: monday = Mon))
  corresponds to a value in the array_datas, the array of school days is incremented (by pushing), with the array of dates at position[i]
  b) if the teacher has more than one class on the same day, the process is repeated.
  */

const onOpen=(e)=> {


  const calculated_dates = [];
  const holidays = []
  const Saturday_work_day = [];
  const StartDay=SpreadsheetApp.getActiveSheet().getRange("i6").getValue();
  const StartMonth=SpreadsheetApp.getActiveSheet().getRange("eq2").getValue();
  const Year=SpreadsheetApp.getActiveSheet().getRange("k6").getValue();
  const Monday = SpreadsheetApp.getActiveSheet().getRange("A3").getValue();
  const Tuesday = SpreadsheetApp.getActiveSheet().getRange("b3").getValue();
  const Wednesday = SpreadsheetApp.getActiveSheet().getRange("c3").getValue();
  const Thursday = SpreadsheetApp.getActiveSheet().getRange("d3").getValue();
  const Friday = SpreadsheetApp.getActiveSheet().getRange("e3").getValue();
  const Saturday = SpreadsheetApp.getActiveSheet().getRange("f3").getValue();
  const Monday_class_2 = SpreadsheetApp.getActiveSheet().getRange("A5").getValue();
  const Tuesday_class_2 = SpreadsheetApp.getActiveSheet().getRange("b5").getValue();
  const Wednesday_class_2 = SpreadsheetApp.getActiveSheet().getRange("c5").getValue();
  const Thursday_class_2 = SpreadsheetApp.getActiveSheet().getRange("d5").getValue();
  const Friday_class_2 = SpreadsheetApp.getActiveSheet().getRange("e5").getValue();
  const Saturday_class_2 = SpreadsheetApp.getActiveSheet().getRange("f5").getValue();


  const GetHolidays = () => {
    const getholidays = SpreadsheetApp.getActiveSheet().getRange("c12:c50").getValues()
    for (let i in getholidays) {
      if(getholidays[i]!=''){
        holidays.push(new Date(getholidays[i]).toString().substring(0, 15));
        
      } else{
        break
      }
        
    }
    return holidays
}




const GetSaturdayWork = () => {
    const getSaturday = SpreadsheetApp.getActiveSheet().getRange("f12:f50").getValues()
    for (let i in getSaturday) {
      if(getSaturday[i]!=''){
        Saturday_work_day.push(new Date(getSaturday[i]).toString().substring(0, 15));
      } else{
        break
      }
        
    }
    return Saturday_work_day
}



  
  const Calendar = (startDate, endDate) => {
    const dates = [];
    let currentDate = startDate;

    function addDays(Days) {
      const date = new Date(this.valueOf());
      date.setDate(date.getDate() + Days);
      return date;
    }
    while (currentDate <= endDate) {
      dates.push(currentDate);
      currentDate = addDays.call(currentDate, 1);
    }
    return dates;
  };


  const Calculate_Dates = (Saturdays) => {
    const dates = Calendar(new Date(Year, StartMonth, StartDay), new Date(Year, 11, 22));

    for (let i in dates) {
        calculated_dates.push(new Date(dates[i]).toString().substring(0, 15));
    }

    let insert_saturdays = calculated_dates.filter(function (dia) {
        if (dia.indexOf("Sat") > -1 || dia.indexOf("Sun") > -1) {
            if (Saturdays.indexOf(dia) > -1) {
                return dia;
            }
        } else {
            return dia;
        }
    });

    const remove_holidays = insert_saturdays.filter(function (dt) {
        return GetHolidays().indexOf(dt) < 0;
    });

    return remove_holidays;
};


  const Search_days_of_class = () => {
    let arr_dates = Calculate_Dates(GetSaturdayWork());
    let arr_days_schools = [];

    for (i in arr_dates) {
        if (
            (Monday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Mon") ||
            (Tuesday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Tue") ||
            (Wednesday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Wed") ||
            (Thursday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Thu") ||
            (Friday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Fri") ||
            (Saturday === "x" &&
                arr_dates[i].toString().substring(0, 3) === "Sat")
        ) {
            arr_days_schools.push(arr_dates[i]);

            if (
                (Monday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Mon") ||
                (Tuesday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Tue") ||
                (Wednesday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Wed") ||
                (Thursday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Thu") ||
                (Friday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Fri") ||
                (Saturday_class_2 === "x" &&
                    arr_dates[i].toString().substring(0, 3) === "Sat")
            ) {
                arr_days_schools.push(arr_dates[i]);
            }
        }
    }

    return arr_days_schools;
};

  let array = Search_days_of_class();

  array.forEach((item, index, array) => (array[index] = [item]));

  console.log(array);

  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(12, 1, array.length, 1).setValues(array);
}

function Clear_data() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const clean = sheet.getRange("a12:a864").clearContent();
  return clean;
}
