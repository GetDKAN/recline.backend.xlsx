;(function($) {
  "use strict";

  $(document).on("ready", function(){

    var backend = {
      backend: "xlsx",
      url: "data/example.xlsx",
      sheet: "apollo-parsed-1737-325_0"
    };

    Excel.fetch(backend).done(function(data){
      console.log(data);
    });
  });
})(jQuery);
