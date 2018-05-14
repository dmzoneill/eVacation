//if depart date not defined, grab the date from the relevant input and turn it into US format for the calendar, then open pop-up calendar
function dateSort(strformname, strinputname, datfirstdate, strcompareinputname) {

    //get an array of the entered date using split()
    var date_array = eval('document.' + strformname + '.' + strinputname + '.value.split(" ")');

    dateField = eval('document.' + strformname + '.' + strinputname);
    dateFieldVal = date_array[0] + ' ' + date_array[1] + ' ' + date_array[2];
    dateMinimum = datfirstdate;
    if (!strcompareinputname == '') dateCompareField = eval('document.' + strformname + '.' + strcompareinputname + '.value')
    else dateCompareField = '';


    //[MFILLAST 08-2006] change windows size + allow resize//	popUpWindow("/eVacation/common/calendar.htm",'cal','width=165,height=215,,left=420,top=100') ;
    popUpWindow("common/calendar.htm", 'cal', 'width=230,height=280,resizable=1,left=420,top=100');

}

function payrollreport(strformname, strstartdatefieldname, strenddatefieldname, strstatusfieldname, strsitefieldname) {
    var datStartDate = eval('document.' + strformname + '.' + strstartdatefieldname + '.value');
    var datEndDate = eval('document.' + strformname + '.' + strenddatefieldname + '.value');
    var status = eval('document.' + strformname + '.' + strstatusfieldname + '.value');
    var site = eval('document.' + strformname + '.' + strsitefieldname + '.value');
    var winURL = "common/reports/payrollreportlite2.asp?datStartDate=" + escape(datStartDate) + '&datEndDate=' + escape(datEndDate) + '&status=' + escape(status) + '&site=' + escape(site) ;
    $.colorbox(
    {
        href: winURL
    });
}

//[MFILLAST 08-2006]
// send the query to the database and display the result in a popup window
// strformname : name of the form with the query
// strqueryfieldname : name of the text area containing the query
function executeQuery(strformname, strqueryfieldname) {
    var query = eval('document.' + strformname + '.' + strqueryfieldname + '.value');
    var winURL = "queryDisplay.asp?query=" + escape(query);
    $.colorbox(
    {
        href: winURL
    });
}


function hrreport(strformname, strmonthfieldname, stryearfieldname, strexemptfieldname, strsitefieldname, strstatusfieldname, stremployeefieldname, strmanagerfieldname) {
    var datStartDate = eval('document.' + strformname + '.' + strmonthfieldname + '.value');
    var datEndDate = eval('document.' + strformname + '.' + stryearfieldname + '.value');
    var strExempt = eval('document.' + strformname + '.' + strexemptfieldname + '.value');
    var strSite = eval('document.' + strformname + '.' + strsitefieldname + '.value');
    var strStatus = eval('document.' + strformname + '.' + strstatusfieldname + '.value');
    var strEmployee = eval('document.' + strformname + '.' + stremployeefieldname + '.value');
    var strManager = eval('document.' + strformname + '.' + strmanagerfieldname + '.value');
    var winURL = "common/reports/hrreport.asp?datMonth=" + escape(datStartDate) + '&datYear=' + escape(datEndDate) + '&status=' + escape(strStatus) + '&site=' + escape(strSite) + '&exempt=' + escape(strExempt) + '&employee=' + escape(strEmployee) + '&manager=' + escape(strManager);
    $.colorbox(
    {
        href: winURL
    });
}

(function ($) {
    $.widget("custom.combobox",
    {
        _create: function () {
            this.wrapper = $("<span>")
                .addClass("custom-combobox")
                .insertAfter(this.element);

            this.element.hide();
            this._createAutocomplete();
            this._createShowAllButton();
        },

        _createAutocomplete: function () {
            var selected = this.element.children(":selected"),
                value = selected.val() ? selected.text() : "";

            this.input = $("<input>")
                .appendTo(this.wrapper)
                .val(value)
                .attr("title", "")
                .addClass("custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left")
                .autocomplete(
                {
                    delay: 0,
                    minLength: 0,
                    source: $.proxy(this, "_source")
                })
                .tooltip(
                {
                    tooltipClass: "ui-state-highlight"
                });

            this._on(this.input,
            {
                autocompleteselect: function (event, ui) {
                    ui.item.option.selected = true;
                    this._trigger("select", event,
                    {
                        item: ui.item.option
                    });
                },

                autocompletechange: "_removeIfInvalid"
            });
        },

        _createShowAllButton: function () {
            var input = this.input,
                wasOpen = false;

            $("<a>")
                .attr("tabIndex", -1)
                .attr("title", "Show All Items")
                .tooltip()
                .appendTo(this.wrapper)
                .button(
                {
                    icons:
                    {
                        primary: "ui-icon-triangle-1-s"
                    },
                    text: false
                })
                .removeClass("ui-corner-all")
                .addClass("custom-combobox-toggle ui-corner-right")
                .mousedown(function () {
                    wasOpen = input.autocomplete("widget").is(":visible");
                })
                .click(function () {
                    input.focus();

                    // Close if already visible
                    if (wasOpen) {
                        return;
                    }

                    // Pass empty string as value to search for, displaying all results
                    input.autocomplete("search", "");
                });
        },

        _source: function (request, response) {
            var matcher = new RegExp($.ui.autocomplete.escapeRegex(request.term), "i");
            response(this.element.children("option").map(function () {
                var text = $(this).text();
                if (this.value && (!request.term || matcher.test(text)))
                    return {
                        label: text,
                        value: text,
                        option: this
                    };
            }));
        },

        _removeIfInvalid: function (event, ui) {

            // Selected an item, nothing to do
            if (ui.item) {
                return;
            }

            // Search for a match (case-insensitive)
            var value = this.input.val(),
                valueLowerCase = value.toLowerCase(),
                valid = false;
            this.element.children("option").each(function () {
                if ($(this).text().toLowerCase() === valueLowerCase) {
                    this.selected = valid = true;
                    return false;
                }
            });

            // Found a match, nothing to do
            if (valid) {
                return;
            }

            // Remove invalid value
            this.input
                .val("")
                .attr("title", value + " didn't match any item")
                .tooltip("open");
            this.element.val("");
            this._delay(function () {
                this.input.tooltip("close").attr("title", "");
            }, 2500);
            this.input.autocomplete("instance").term = "";
        },

        _destroy: function () {
            this.wrapper.remove();
            this.element.show();
        }
    });
})(jQuery);

function setCookie(val) 
{
    var d = new Date();
    d.setTime(d.getTime() + (30*60*1000));
    var expires = "expires="+d.toUTCString();
    document.cookie = "evac=" + val + "; " + expires;
}

function getCookie() 
{
  var name = "evac=";
  var ca = document.cookie.split(';');
  for(var i=0; i<ca.length; i++) 
  {
    var c = ca[i];
    while (c.charAt(0)==' ') c = c.substring(1);
    if (c.indexOf(name) == 0) return c.substring(name.length,c.length);
  }
  return "";
}


function validatewwid()
{
  
}


$(document).ready(function () 
{
    $( '#userwwid' ).numeric();
    
    $( '#submitwwid' ).click( function()
    {
      if( $( '#userwwid' ).val().length == 8 )
      {
        $( '#addwwidfeedback').html('');
        $( '#addwwidform' ).submit();
      }
      else
      {
        $( '#addwwidfeedback').html('WWID should be 8 digits');
      }
    });
    
    $("#showOtherLeave").click(function () 
    {
      $("#otherLeaveTable").show("slow");
    })

    $(".evdatepicker").datepicker(
    {
        showOn: "button",
        buttonImage: "common/images/calendar.gif",
        buttonImageOnly: true,
        buttonText: "Select date",
        numberOfMonths: 2,
        showButtonPanel: true
    });

    $(".evdatepicker").datepicker("option", "dateFormat", "d M yy");

    $('input[type="submit"]').button();
    $('input[type="reset"]').button();

    $('input:text, input:password').button().css(
    {
        'font': 'inherit',
        'color': 'inherit',
        'text-align': 'left',
        'outline': 'none',
        'cursor': 'text'
    });

    $("input[type=checkbox]").switchButton(
    {
        width: 80,
        height: 20,
        button_width: 55
    });

    $(".basicselect").selectmenu();
    $(".ampmselect").selectmenu();

    $(".wwidcombobox").combobox();

    $("#yearselect").selectmenu();
    $("#yearselect").on("selectmenuchange", function (event, ui) {
        document.frmYearSelector.submit();
    });

    $(".iframe").colorbox(
    {
        iframe: true,
        width: "80%",
        height: "80%"
    });

    $("#dialog-message").dialog(
    {
        autoOpen: false,
        height: 350,
        width: 450,
        modal: true,
        buttons:
        {
            Ok: function () {
                $(this).dialog("close");
            }
        }
    });

    $("#errorshowlink").click(function () {
        $("#dialog-message").dialog("open");
    });
	
	var cval = getCookie();
		
	if( cval == 1 )
		return;
	
	if( cval == 2 )
	{
		document.location.href = "https://shannon.intel.com/eVacationOld/default.asp";
		return;
	}
})