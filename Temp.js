function myFunction() {
      for (var i = 0; i < guestList.length; i++) {
      var guest = guestList[i];

      var dsfa = guest.getEmail();
      var x1 = guest.getGuestStatus() === CalendarApp.GuestStatus.YES;

      var guestStatus = guest.getGuestStatus()
      var yes = guestStatus.YES.toString();
      console.log("GuestStatus")
    }





        // var formatEnd=Utilities.formatDate(end, 'Europe/Warsaw', 'HH-mm-ss');
    // console.log(formatEnd);
    // if (formatEnd=="00-00-00"){
    //   continue
    // }
}
