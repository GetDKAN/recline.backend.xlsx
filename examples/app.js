;(function($) {
  "use strict";

  $(document).on("ready", function(){

    var backend = {
      backend: "xlsx",
      url: "data/testxls1.xlsx",
    };

    Excel.fetch(backend).done(function(data){
      console.log(data);
    });
  });
})(jQuery);
