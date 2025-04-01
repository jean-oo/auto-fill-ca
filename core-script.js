// ==UserScript==
// @name         login page
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  auto fill in login page
// @author       Jean
// @match        https://prson-srpel.apps.cic.gc.ca/en/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=tampermonkey.net
// @grant        none
// @resource     excel  /Users/xueqin/Documents/test\ sample.xlsx
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.5/xlsx.full.min.js
// @require      https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
// @require      https://cdn.bootcss.com/xlsx/0.15.6/xlsx.core.min.js
// @require      https://cdn.bootcss.com/jquery/3.1.1/jquery.min.js
// @require      https://code.jquery.com/jquery-2.1.4.min.js

// ==/UserScript==
function simulateMouseClick(targetNode) {
  function triggerMouseEvent(targetNode, eventType) {
      var clickEvent = document.createEvent('MouseEvents');
      clickEvent.initEvent(eventType, true, true);
      targetNode.dispatchEvent(clickEvent);
  }
  ["mouseover", "mousedown", "mouseup", "click"].forEach(function(eventType) {
      triggerMouseEvent(targetNode, eventType);
  });
}

(function() {
  'use strict';
  var $$ = jQuery.noConflict();

  /* globals jQuery, $, waitForKeyElements */
  $(function(){
      var datas = new Array();
      var index = 12;
      var workbook = null;
      var sheet_name_list = null;
      var dependentInd0008 = 55;
      var form0008 = null;
      var dependent0008Indices = null;
      var depPersonalDetailIndices = null;
      var depEduIndices = null;
      var depLanguageIndices = null;
      var depPassportIndices = null;
      var depNationalIDIndices = null;

      var form5406 = null;
      var form5562 = null;
      var sectionAIndices = null;
      var sectionBIndices = null;
      var sectionCIndices = null;
      var sectionAIndices5562 = null;
      var sectionBIndices5562 = null;
      var sectionCIndices5562 = null;
      var depNum = null;
      var form0008_counter = 0;
      var form5406_counter = 0;
      var form5562_counter = 0;

      var form5669 = null;
      var form5569_counter = 0;
      var personalDetailIndices = null;
      var QuestionnaireIndices = null;
      var educationIndices = null;
      var persHistIndices = null;
      var membershipIndices = null;
      var govermentIndices = null;
      var militaryIndices = null;
      var addresseIndices = null;

      var changeEvent = new Event('change', {bubbles:true});



      //add file input
      var input_html = `
  <div id='control' style='background-color: white; width: 300px; height: 200px; position: fixed; z-index: 2; left: 80%; top: 5%'>
     <input type='file' id='xlsxfile' style='width: 200px; height: 25px' />
     <style>
     /* Add font-size style for the buttons */
     #start, #profile_page, #addDep, #fill0008, #personalDetail0008,
     #contact0008, #passport0008, #nationalID, #edu_occup, #language0008,
     #dep0008-1, #dep0008-2, #dep0008-3, #dep0008-4, #dep0008-5, #next0008, #prev0008,
     #sectionA5406, #sectionB5406, #sectionC5406, #nextDependant, #prevDependant,
     #persDetails5669, #Quesitonaire5669, #Education5669, #Edu-add5669,
     #persHist5669, #membership5669, #govPosition5669, #military5669,
     #address5669, #prev5669, #next5669, #sectionA5562, #sectionB5562, #sectionC5562 {
         font-size: 14px; /* Set your desired font size */
         background-color: #000000;
         }
     #start:hover { background-color: #ffd700; color: #000; }
     #profile_page:hover { background-color: #ffd700; color: #000; }
     #addDep:hover { background-color: #ffd700; color: #000; }
     #fill0008:hover { background-color: #ffd700; color: #000; }
     #personalDetail0008:hover { background-color: #ffd700; color: #000; }
     #contact0008:hover { background-color: #ffd700; color: #000; }
     #passport0008:hover { background-color: #ffd700; color: #000; }
     #nationalID:hover { background-color: #ffd700; color: #000; }
     #edu_occup:hover { background-color: #ffd700; color: #000; }
     #language0008:hover { background-color: #ffd700; color: #000; }
     #dep0008-1:hover { background-color: #ffd700; color: #000; }
     #dep0008-2:hover { background-color: #ffd700; color: #000; }
     #dep0008-3:hover { background-color: #ffd700; color: #000; }
     #dep0008-4:hover { background-color: #ffd700; color: #000; }
     #dep0008-5:hover { background-color: #ffd700; color: #000; }
     #next0008:hover { background-color: #ffd700; color: #000; }
     #prev0008:hover { background-color: #ffd700; color: #000; }
     #sectionA5406:hover { background-color: #ffd700; color: #000; }
     #sectionB5406:hover { background-color: #ffd700; color: #000; }
     #sectionC5406:hover { background-color: #ffd700; color: #000; }
     #nextDependant:hover { background-color: #ffd700; color: #000; }
     #prevDependant:hover { background-color: #ffd700; color: #000; }
     #persDetails5669:hover { background-color: #ffd700; color: #000; }
     #Quesitonaire5669:hover { background-color: #ffd700; color: #000; }
     #Education5669:hover { background-color: #ffd700; color: #000; }
     #Edu-add5669:hover { background-color: #ffd700; color: #000; }
     #persHist5669:hover { background-color: #ffd700; color: #000; }
     #membership5669:hover { background-color: #ffd700; color: #000; }
     #govPosition5669:hover { background-color: #ffd700; color: #000; }
     #military5669:hover { background-color: #ffd700; color: #000; }
     #address5669:hover { background-color: #ffd700; color: #000; }
     #next5669:hover { background-color: #ffd700; color: #000; }
     #prev5669:hover { background-color: #ffd700; color: #000; }
     #sectionA5562:hover { background-color: #ffd700; color: #000; }
     #sectionB5562:hover { background-color: #ffd700; color: #000; }
     #sectionC5562:hover { background-color: #ffd700; color: #000; }


   </style>
      <div>
      <button type='button' id='start' style='width: 120px; height: 25px; margin-left: 90px; padding: 0'>开始</button>
      </div>
      <div>
      <button id='profile_page' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Profile Page</button>
      <!-- <button id='residentialAdd' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Residential Address</button> -->
      <button id='addDep' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>AddDependent</button>
      </div>
      <div>
      <button type='button' id='form0008' style='width: 80px; height: 30px; margin-left: 112px; padding: 0'>0008</button>
      <button id='fill0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>ApplicationDetails</button>
      <button id='personalDetail0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>PersonalDetails</button>
      <button id='contact0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>ContactInformation</button>
      <button id='passport0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Passport</button>
      <button id='nationalID' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>National_ID</button>
      <button id='edu_occup' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Edu/OccupDetails</button>
      <button id='language0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>LanguageDetail</button>
      <button id='dep0008-1' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Dependant-1</button>
      <button id='dep0008-2' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Dependant-2</button>
      <button id='dep0008-3' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Dependant-3</button>
      <button id='dep0008-4' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Dependant-4</button>
      <button id='dep0008-5' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Dependant-5</button>
      <button id='prev0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>PreviousDependant</button>
      <button id='next0008' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>NextDependant</button>
      </div>
      <div>
      <button type='button' id='form5406' style='width: 80px; height: 30px; margin-left: 112px; padding: 0'>5406</button>
      <button id='sectionA5406' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionA</button>
      <button id='sectionB5406' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionB</button>
      <button id='sectionC5406' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionC</button>
      <button id='nextDependant' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>NextDependant</button>
      <button id='prevDependant' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding:0'>PreviousDependant</button>
      </div>
      <div>
      <button type='button' id='form5669' style='width: 80px; height: 30px; margin-left: 112px; padding: 0'>5669</button>
      <button id='persDetails5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>PersonalDetails</button>
      <button id='Quesitonaire5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Quesitonaire</button>
      <button id='Education5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Education</button>
      <button id='Edu-add5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>EducationHistory</button>
      <button id='persHist5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>PersonalHistory</button>
      <button id='membership5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Membership</button>
      <button id='govPosition5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Government</button>
      <button id='military5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Military</button>
      <button id='address5669' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>Address</button>
      <button id='next5669' type='button' style='width: 120px; height: 25px; padding: 0; margin-left: 20px'>NextPerson</button>
      <button id='prev5669' type='button' style='width: 120px; height: 25px; padding: 0; margin-left: 20px'>PrevPerson</button>
      </div>
      <div>
      <button type='button' id='form5562' style='width: 80px; height: 30px; margin-left: 112px; padding: 0'>5562</button>
      <button id='sectionA5562' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionA</button>
      <button id='sectionB5562' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionB</button>
      <button id='sectionC5562' type='button' style='width: 120px; height: 25px; margin-left: 20px; padding: 0'>SectionC</button><br>
      <style>
</style>
      </div>
      <textarea id="texta" name="w3review" rows="1" cols="25"> </textarea>
  </div>
`;



      $$('body').prepend(input_html);

      //upload file
      $$('#xlsxfile').change(function(){
          var reader = new FileReader();
          reader.readAsBinaryString(document.getElementById("xlsxfile").files[0]);
          reader.onloadend = function (evt) {
              if(evt.target.readyState == FileReader.DONE){
                  var data = reader.result;
                  workbook = XLSX.read(data, { type: 'binary' });
              }
              sheet_name_list = workbook.SheetNames;
              datas = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {header:1});
              $$('#texta').append(datas);
              var addDepPart = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]], {header:1});
              depNum = addDepPart[11][1];
              form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
              dependent0008Indices = findIndicesOfValue(form0008, "Dependants");
              console.log("dependent0008Indices"+dependent0008Indices);
              depPersonalDetailIndices = findIndicesOfValue(form0008, "Dep Personal details");
              console.log("depPersonalDetailIndices"+depPersonalDetailIndices);
              depEduIndices = findIndicesOfValue(form0008, "Dep Education/occupation detail");
              depLanguageIndices = findIndicesOfValue(form0008, "Dep Language detail");
              depPassportIndices = findIndicesOfValue(form0008, "Dep Passport");
              depNationalIDIndices = findIndicesOfValue(form0008, "Dep National identity document");
              form5406 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[3]], {header:1});
              sectionAIndices = findIndicesOfValue(form5406, "Section A");
              sectionBIndices = findIndicesOfValue(form5406, "Section B： Children");
              sectionCIndices = findIndicesOfValue(form5406, "Section C: Siblings");
              form5562 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[5]], {header:1});
              sectionAIndices5562 = findIndicesOfValue(form5562, "Section A: Principal applicant (yourself)");
              sectionBIndices5562 = findIndicesOfValue(form5562, "Section B: Your spouse or common-law partner");
              sectionCIndices5562 = findIndicesOfValue(form5562, "Section C: Your dependant children 18 years or older");
              form5669 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[4]], {header:1});
              personalDetailIndices = findIndicesOfValue(form5669, "Personal Detail");
              QuestionnaireIndices = findIndicesOfValue(form5669, "Questionnaire");
              educationIndices = findIndicesOfValue(form5669, "Education");
              persHistIndices = findIndicesOfValue(form5669, "Personal History");
              membershipIndices = findIndicesOfValue(form5669, "Membership and association with organizations");
              govermentIndices = findIndicesOfValue(form5669, "Government positions");
              militaryIndices = findIndicesOfValue(form5669, "Military and paramilitary service");
              addresseIndices = findIndicesOfValue(form5669, "Addresses");

          }
          alert('file upload success');
      });

      function imitateKeyInput(el, keyChar){
          if (el) {
              const keyEventInit = {bubble:false, cancelable:false, composed:false,key:'', code:'',location:0};
              el.focus();
              document.execCommand('insertText', false, keyChar);
              console.log("keyChar", keyChar);
              el.dispatchEvent(new Event('change', {bubbles:true}));
              //}
          } else {
              console.log("el is null");
          }
      }
      async function delay(ms) {
          return new Promise(resolve => setTimeout(resolve, ms));
      }

      // profile page
      $$("#profile_page").click(async function(){
          console.log("next");
          settingProfile()
          await delay(500);
          settingProfileResAdd();
      })
      //$$("#residentialAdd").click(function(){
      //    settingProfileResAdd()
      //})

      function settingProfile(){
          var sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]], {header:1});
          var preferredLanguage = document.querySelectorAll('select[id = "profileForm-correspondence"]')[0];
          preferredLanguage.selectedIndex = 1;
          var event = new Event("change");
          preferredLanguage.dispatchEvent(event);
          const fName= document.querySelector('#profileForm-familyName');
          const gName = document.querySelector('#personalDetailsForm-givenName');
          const birth = document.querySelector('#personalDetailsForm-dob');
          const pobox = document.querySelector('#postOfficeBox');
          const aptUnit = document.querySelector('#apartmentUnit');
          const strNum = document.querySelector('#streetNumber');
          const strName = document.querySelector('#streetName');
          const city = document.querySelector('#city');
          var country = document.querySelectorAll('select[id = "country"]')[0];
          country.selectedIndex = 42;
          country.dispatchEvent(event);
          var province = document.querySelector('#province');
          window.setTimeout(()=>{
              province.selectedIndex = 2;
              province.dispatchEvent(event);
          },200);
          const postalcode = document.querySelector('#postalCode');
          const district = document.querySelector('#district');

          imitateKeyInput(fName, sheet2[1][0]);
          imitateKeyInput(gName, sheet2[1][1]);
          imitateKeyInput(birth, sheet2[1][2]);
          imitateKeyInput(pobox, sheet2[2][0]);
          imitateKeyInput(aptUnit, sheet2[2][1]);
          imitateKeyInput(strNum, sheet2[2][2]);
          imitateKeyInput(strName, sheet2[2][3]);
          imitateKeyInput(city, sheet2[3][0]);
          imitateKeyInput(country, sheet2[3][1]);
          imitateKeyInput(province, sheet2[3][2]);
          imitateKeyInput(postalcode, sheet2[4][0]);
      }
      function settingProfileResAdd(){
          var sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]], {header:1});
          if(sheet2[5][0] === "Yes"){
              const sameAdd = document.querySelector('#yes');
          sameAdd.checked = true;
          sameAdd.dispatchEvent(new Event('change', {bubbles:true}));
              return;
          }

          const notSameAdd = document.querySelector('#no');
          notSameAdd.checked = true;
          notSameAdd.dispatchEvent(new Event('change', {bubbles:true}));

          //var sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]], {header:1});
          const resAptUnit = document.querySelectorAll('input[id="apartmentUnit"]')[1];
          const resStrNum = document.querySelectorAll('input[id="streetNumber"]')[1];
          const resStrName = document.querySelectorAll('input[id="streetName"]')[1];
          const resCity = document.querySelectorAll('input[id="city"]')[1];
          var resCountry = document.querySelectorAll('select[id = "country"]')[1];
          console.log("resCity"+resCountry);
          console.log("resCountry"+resCountry);
          var resCountryFromExecl = sheet2[8][1];
          var countryIndex = 0;
          if (resCountryFromExecl === "Canada"){
              countryIndex = 42;
          }else if(resCountryFromExecl === "China"){
              countryIndex = 49;
          }
          resCountry.selectedIndex = countryIndex;
          resCountry.dispatchEvent(new Event('change', {bubbles:true}));
          const resProvince = document.querySelectorAll('input[id="province"]')[1];
          const resPostalcode = document.querySelectorAll('input[id="postalCode"]')[1];
          const resDistrict = document.querySelectorAll('input[id="district"]')[1];

          imitateKeyInput(resAptUnit, sheet2[7][0]);
          imitateKeyInput(resStrNum, sheet2[7][1]);
          imitateKeyInput(resStrName, sheet2[7][2]);
          imitateKeyInput(resCity, sheet2[8][0]);
          // imitateKeyInput(resCountry, sheet2[8][1]);
          imitateKeyInput(resProvince, sheet2[8][2]);
          imitateKeyInput(resPostalcode, sheet2[9][0]);
          console.log("resPostalcode 9 0 "+sheet2[9][0]);
          if(countryIndex !==42){
              imitateKeyInput(resDistrict, sheet2[9][1]);
              console.log("resDistrict 9 1 "+sheet2[9][1]);
          }
      }

      //start
      $$('#start').click(function(){
          //setting(datas[index]);
          const el_pass= document.querySelector('#password');
          const el_user = document.querySelector('#username');
          imitateKeyInput(el_user, datas[0][0]);
          console.log(datas[0][0]);
          imitateKeyInput(el_pass, datas[0][1]);
          var signInButton = document.querySelector('button[category="primary"][type="submit"]')
          signInButton.removeAttribute('disabled')
          //document.querySelector('form[autocomplete="off"]').setAttribute('autocomplete', 'on');
          signInButton.click();
          window.addEventListener('load', function(){
              console.log("console");
          });
      });

      //add dependent
      $$("#addDep").click(function(){
          var addDepPart = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]], {header:1});
          // index = 12;
          var depNum = addDepPart[11][1];
          // for(let i = 0; i<depNum; i++){
          fillDependent(addDepPart);
          // }
      })

      function fillDependent(sheet2){
          //$$(document).ready(function() {
          //  window.scrollTo({ top: 500, behavior: 'smooth' });
          //window.setTimeout(()=>{
          // yes/no select code
          const addButton = document.querySelectorAll('button[category="secondary"]')[2];
          addButton.click();
          addButton.dispatchEvent(new Event("click", {bubbles:true}));

          //window.setTimeout(()=>{
          var data = sheet2[index++][0];

          const yesInput = document.querySelector('.form-radio__option#dependantYes');
          const noInput = document.querySelector('.form-radio__option#dependantNo');

          if (data === "Yes") {
              // Select "yes" option
              yesInput.checked = true;
              yesInput.dispatchEvent(new Event('change', {bubbles:true}));
          } else if (data === "No") {
              // Select "no" option
              noInput.checked = true;
              noInput.dispatchEvent(new Event('change', {bubbles:true}));
          }
          //  },1000)
          //}, 5000)

          var program_dropdown = document.querySelectorAll('select[id = "dependantForm-relationship"]')[0];
          var relationship = sheet2[index++][0];
          var relaList = ["Adopted Child","Child","Common-Law Partner","Grandchild","Other","Spouse","Step-Child","Step-Grandchild","Parent","Adoptive Parent"]
          program_dropdown.selectedIndex = relaList.indexOf(relationship)+1;
          var event = new Event("change");
          program_dropdown.dispatchEvent(event);

          const depFName = document.querySelector('input[id="dependantForm-familyName"]');
          const depGName = document.querySelector('input[id="dependantForm-givenName"]');
          const dob = document.querySelectorAll('input[id="personalDetailsForm-dob"]')[1];
          imitateKeyInput(depFName, sheet2[index++][0]);
          imitateKeyInput(depGName, sheet2[index++][0]);
          var date = sheet2[index++][0];
          window.setTimeout(()=>{
              imitateKeyInput(dob, date);
          },500);
          const saveDetails=document.querySelector('button[category = "primary"][type="button"]')
          saveDetails.click();
          //})
      }

      $$("#fill0008").click(function(){
          fill0008()
      })

      /* Fill 0008*/
      function fill0008(){
          const event = new Event("change");
          //var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          changeEvent = new Event('change', {bubbles:true});
          const interview = document.querySelector('select[id="interview"]');
          interview.selectedIndex = 123;
          interview.dispatchEvent(changeEvent);
          const interpreter = document.querySelector('select[id="interpreterRequested"]');
          interpreter.selectedIndex = 2;
          interpreter.dispatchEvent(changeEvent);
          var provs = ["AB","BC","MB","NB","NL","NS","NT","NU","ON","PE","QC","SK","YT"]
          var prov = document.querySelector('select[id="province"]');
          var city = document.querySelector('select[id="city"]');
          prov.selectedIndex = provs.indexOf(form0008[3][2])+1;
          prov.dispatchEvent(event);
          city.selectedIndex = 215;
          city.dispatchEvent(event);
      }
      $$("#personalDetail0008").click(function(){
          personalDetails0008(6,2);
      });

      async function personalDetails0008(row,col){
          var cellCol = col;
          var cellRow = row;
          //var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          var hasNickName = form0008[cellRow][cellCol];
          const yesNickName = document.querySelector('#personalDetailsForm-usedOtherName-yes');
          const noNickName = document.querySelector('#personalDetailsForm-usedOtherName-no');
          if (hasNickName === "Yes") {
              yesNickName.checked = true;
              yesNickName.dispatchEvent(new Event('change', {bubbles:true}));
              await(delay(500));
              var nickFName = document.querySelector('#personalDetailsForm-otherFamilyName');
              var nickGName = document.querySelector('#personalDetailsForm-otherGivenName');
              imitateKeyInput(nickFName, form0008[cellRow][++cellCol]);
              imitateKeyInput(nickGName, form0008[cellRow][++cellCol]);
              return new Promise(async(resolve)=>{
                  resolve(await personalDetails0008afterUCI(cellRow,col))
              });

          } else if (hasNickName === "No") {
              // Select "no" option
              noNickName.checked = true;
              noNickName.dispatchEvent(new Event('change', {bubbles:true}));
              return new Promise(async(resolve)=>{
                  resolve(await personalDetails0008afterUCI(cellRow,col))
              });
          }
      }


      async function personalDetails0008afterUCI(row,col){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          var cellRow = row;
          var cellCol = col;
          var UCI = document.querySelector('#personalDetailsForm-uci');
          imitateKeyInput(UCI, form0008[++cellRow][col]);
          var sex = document.querySelector('#personalDetailsForm-sex');
          var sexInput = form0008[++cellRow][col];
          sex.selectedIndex = sexInput;
          sex.dispatchEvent(new Event('change', {bubbles:true}));
          var eyeColor = document.querySelector('#personalDetailsForm-eyeColour');
          var eyeColorInput = form0008[++cellRow][col];
          eyeColor.selectedIndex = eyeColorInput;
          eyeColor.dispatchEvent(new Event('change', {bubbles:true}));
          var height = document.querySelector('#personalDetailsForm-heightInCM');
          imitateKeyInput(height, form0008[++cellRow][col]);
          var placeOfBirth = document.querySelector('#personalDetailsForm-cityOfBirth');
          imitateKeyInput(placeOfBirth, form0008[++cellRow][col]);
          var countryOfBirth = document.querySelector('#personalDetailsForm-countryOfBirth');
          countryOfBirth.selectedIndex = 49;
          countryOfBirth.dispatchEvent(new Event('change', {bubbles:true}));
          var citiCountry = document.querySelector('#personalDetailsForm-citizenship1');
          citiCountry.selectedIndex = 37;
          citiCountry.dispatchEvent(new Event('change', {bubbles:true}));
          var resCountry = document.querySelector('#personalDetailsForm-currentCountry');
          cellRow += 3;
          var resCountryForm = form0008[cellRow][col];
          resCountry.selectedIndex = resCountryForm;
          resCountry.dispatchEvent(new Event('change', {bubbles:true}));
          var status = document.querySelector('#personalDetailsForm-immigrationStatus');
          var statusForm = form0008[++cellRow][col];
          status.selectedIndex = statusForm;
          status.dispatchEvent(new Event('change', {bubbles:true}));
          if (statusForm ===3 || statusForm ===4 ||statusForm ===5){
              await delay(150);
              var resStatuFrom = document.querySelector('#personalDetailsForm-startDateofImmigrationStatus');
              cellCol = col
              imitateKeyInput(resStatuFrom, form0008[cellRow][++cellCol]);
              var resStatuTo = document.querySelector('#personalDetailsForm-endDateOfImmigrationStatus');
              imitateKeyInput(resStatuTo, form0008[cellRow][++cellCol]);
              var dateOfLastEntry = document.querySelector('#personalDetailsForm-dateOfLastEntry');
              imitateKeyInput(dateOfLastEntry, form0008[cellRow][++cellCol]);
              var placeOfLastEntry = document.querySelector('#personalDetailsForm-placeOfLastEntry');
              imitateKeyInput(placeOfLastEntry, form0008[cellRow][++cellCol]);
          }
          const yesMore5 = document.querySelector('#personalDetailsForm-hasPreviousCountries-yes');
          const noMore5 = document.querySelector('#personalDetailsForm-hasPreviousCountries-no');
          cellRow = cellRow+1;
          var more5 = form0008[cellRow][col];

          if (more5 === "Yes") {
              var preCountryIte = 0
              yesMore5.checked = true;
              yesMore5.dispatchEvent(new Event('change', {bubbles:true}));

              cellCol = col;
              while(form0008[cellRow][++cellCol] != undefined){
                  await delay(400);
                  const prevCountry = document.querySelector('#personalDetailsForm-prevCountry'+preCountryIte);
                  prevCountry.selectedIndex= form0008[cellRow][cellCol];
                  prevCountry.dispatchEvent(new Event('change', {bubbles:true}));
                  const prevImmigrationStatus = document.querySelector('#personalDetailsForm-prevImmigrationStatus'+preCountryIte);
                  prevImmigrationStatus.selectedIndex= form0008[cellRow][++cellCol];
                  prevImmigrationStatus.dispatchEvent(new Event('change', {bubbles:true}));
                  const prevStartDateImmigrationStatus = document.querySelector('#personalDetailsForm-prevStartDateOfImmigrationStatus'+preCountryIte);
                  imitateKeyInput(prevStartDateImmigrationStatus, form0008[cellRow][++cellCol]);

                  const prevEndDateImmigrationStatus = document.querySelector('#personalDetailsForm-prevEndDateOfImmigrationStatus'+preCountryIte);
                  imitateKeyInput(prevEndDateImmigrationStatus, form0008[cellRow][++cellCol]);
                  if(form0008[cellRow][cellCol+1] != undefined){
                      document.querySelector('button[category = "primary"][type="button"]').click();
                  }
                  preCountryIte++;
              }

          } else if (more5 === "No") {
              noMore5.checked = true;
              noMore5.dispatchEvent(new Event('change', {bubbles:true}));
          }
          const martial = document.querySelector('#personalDetailsForm-maritalStatus');
          var martialStatu = form0008[++cellRow][col];
          martial.selectedIndex = martialStatu;
          martial.dispatchEvent(new Event('change', {bubbles:true}));
          cellCol = col;

          if(martialStatu === 2 || martialStatu === 4){
              await(delay(500));
              const marriedDate = document.querySelector('#personalDetailsForm-dateOfMarriageOrCommonLaw');
              const spouseFname = document.querySelector('#personalDetailsForm-familyNameOfSpouse');
              const spouseGname = document.querySelector('#personalDetailsForm-givenNameOfSpouse')
              imitateKeyInput(marriedDate, form0008[cellRow][++cellCol]);
              imitateKeyInput(spouseFname, form0008[cellRow][++cellCol]);
              imitateKeyInput(spouseGname, form0008[cellRow][++cellCol]);
          }
          return new Promise(async(resolve)=>{
              resolve(await personalDetails0008afterCurrentMartal(cellRow,col))
          });

      }
      async function personalDetails0008afterCurrentMartal(row,col){
          var cellRow = row;
          var cellCol = col;
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          const yesPreRela = document.querySelector('#personalDetailsForm-previouslyMarriedOrCommonLaw-yes');
          const noPreRela = document.querySelector('#personalDetailsForm-previouslyMarriedOrCommonLaw-no');
          var preRela = form0008[++cellRow][col];
          if (preRela === "Yes") {
              yesPreRela.checked = true;
              yesPreRela.dispatchEvent(new Event('change', {bubbles:true}));
              await(delay(500));
              var spouseFName = document.querySelector('#previousRelationshipForm-previousSpouseFamilyName');
              var spouseGName = document.querySelector('#previousRelationshipForm-previousSpouseGivenName');
              var spouseDOB = document.querySelector('#previousRelationshipForm-previousSpouseDob');
              var spouseRela = document.querySelector('#previousRelationshipForm-typeOfRelationship');
              var startDateRela = document.querySelector('#previousRelationshipForm-startDateofRelationship');
              var endDateRela = document.querySelector('#previousRelationshipForm-endDateOfRelationship');
              imitateKeyInput(spouseFName, form0008[++cellRow][col]);
              imitateKeyInput(spouseGName, form0008[++cellRow][col]);
              imitateKeyInput(spouseDOB, form0008[++cellRow][col]);
              spouseRela.selectedIndex= form0008[++cellRow][col];
              spouseRela.dispatchEvent(new Event('change', {bubbles:true}));
              imitateKeyInput(startDateRela, form0008[++cellRow][col]);
              imitateKeyInput(endDateRela, form0008[++cellRow][col]);

          } else if (preRela === "No") {
              noPreRela.checked = true;
              noPreRela.dispatchEvent(new Event('change', {bubbles:true}));
              cellRow = cellRow + 6;
          }
          return new Promise((resolve)=>{resolve(cellRow)});
      }
      $$('#contact0008').click(function(){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          const primaryNumArea = document.querySelector('#primaryNA');
          primaryNumArea.checked = true;
          primaryNumArea.dispatchEvent(new Event('change', {bubbles:true}));
          const primaryType = document.querySelector('#PrimaryType');
          primaryType.selectedIndex = 2;
          primaryType.dispatchEvent(new Event('change', {bubbles:true}));
          const primaryNum = document.querySelector('#PrimaryNumber');
          console.log("form0008[28][2]",form0008[28][2]);
          imitateKeyInput(primaryNum, form0008[28][2]);

          const altNA = document.querySelector('#altNA');
          const altOther = document.querySelector('#altOther');
          var altType = form0008[29][2];
          // console.log("altType 461" +altType);
          console.log("462"+ altType === "Canada/US");
          if(altType === 1){
              console.log("altType 463" +altType);
              altNA.checked = true;
              altNA.dispatchEvent(new Event('change', {bubbles:true}));
              console.log("Canada/US clicked" );
          } else if(altType === 2){
              altOther.checked = true;
              altOther.dispatchEvent(new Event('change', {bubbles:true}));
          }
          const AlternateType = document.querySelector('#AlternateType');
          AlternateType.selectedIndex = form0008[30][2];
          AlternateType.dispatchEvent(new Event('change', {bubbles:true}));
          const AlternateCountryCode = document.querySelector('#AlternateCountryCode');
          if(altType !=="Canada/US"){
              imitateKeyInput(AlternateCountryCode, form0008[31][2]);
          }
          const AlternateNumber = document.querySelector('#AlternateNumber');
          imitateKeyInput(AlternateNumber, form0008[32][2]);

          const contactYes = document.querySelector('#contactYes');
          contactYes.checked = true;
          contactYes.dispatchEvent(new Event('change', {bubbles:true}));
      });

      async function passport(row,col){
          var cellRow = row;
          var cellCol = col;
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          const yesPassport = document.querySelector('#validPassportYes');
          const noPassport = document.querySelector('#validPassportNo');
          var hasPassport = form0008[cellRow][cellCol];
          if (hasPassport ==="Yes"){
              yesPassport.checked = true;
              yesPassport.dispatchEvent(new Event('change', {bubbles:true}));
          } else if (hasPassport ==="No"){
              noPassport.checked = true;
              noPassport.dispatchEvent(new Event('change', {bubbles:true}));
          }
          await delay(100);
          const passportNumber = document.querySelector('#passportNumber');
          const countryOfIssue = document.querySelector('#countryOfIssue');
          const issueDate = document.querySelector('#issueDate');
          const expiryDate = document.querySelector('#expiryDate');
          imitateKeyInput(passportNumber,"");
          imitateKeyInput(passportNumber,form0008[++cellRow][cellCol]);
          console.log("countryOfIssue "+cellRow + cellCol +" "+form0008[cellRow][cellCol]);
          countryOfIssue.selectedIndex = 37;
          ++cellRow;
          countryOfIssue.dispatchEvent(new Event('change', {bubbles:true}));
          imitateKeyInput(issueDate,form0008[++cellRow][cellCol]);
          imitateKeyInput(expiryDate,form0008[++cellRow][cellCol]);
          return new Promise((resolve)=>{resolve(cellRow)});
      }
      async function nationalID(row,col){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          var cellRow = row;
          var cellCol = col;
          const yesID = document.querySelector('input[id="NICYes"]');
          var noID = document.querySelector('input[id="NICNo"]');
          var hasID = form0008[cellRow][cellCol];
          if (hasID === "Yes"){
              yesID.checked = true;
              yesID.dispatchEvent(new Event('change', {bubbles:true}));
          } else if (hasID === "No"){
              noID.checked = true;
              noID.dispatchEvent(new Event('change', {bubbles:true}));
          }
          await delay(100);
          const nationalIdentityNumber = document.querySelector('#nationalIdentityNumber');
          const IDcountryOfIssue = document.querySelector('#countryOfIssue');
          const IDissueDate = document.querySelector('#issueDate');
          const IDexpiryDate = document.querySelector('#expiryDate');
          imitateKeyInput(nationalIdentityNumber,"");
          imitateKeyInput(nationalIdentityNumber,form0008[++cellRow][cellCol]);
          IDcountryOfIssue.selectedIndex = 37;
          IDcountryOfIssue.dispatchEvent(new Event('change', {bubbles:true}));
          ++cellRow;
          imitateKeyInput(IDissueDate,form0008[++cellRow][cellCol]);
          imitateKeyInput(IDexpiryDate,form0008[++cellRow][cellCol]);
          return new Promise((resolve)=>{resolve(cellRow)});
      };

      $$('#passport0008').click(async function(){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          await passport(35,2);
      });
      $$('#nationalID').click(function(){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          const yesID = document.querySelector('input[id="NICYes"]');
          yesID.checked = true;
          yesID.dispatchEvent(new Event('change', {bubbles:true}));
          window.setTimeout(()=>{
              const nationalIdentityNumber = document.querySelector('#nationalIdentityNumber');
              const IDcountryOfIssue = document.querySelector('#countryOfIssue');
              const IDissueDate = document.querySelector('#issueDate');
              const IDexpiryDate = document.querySelector('#expiryDate');
              imitateKeyInput(nationalIdentityNumber,"");
              imitateKeyInput(nationalIdentityNumber,form0008[42][2]);
              console.log("id" +nationalIdentityNumber,form0008[42][2]);
              IDcountryOfIssue.selectedIndex = 37;
              IDcountryOfIssue.dispatchEvent(new Event('change', {bubbles:true}));
              imitateKeyInput(IDissueDate,form0008[44][2]);
              imitateKeyInput(IDexpiryDate,form0008[45][2]);
          },250);});
      $$('#edu_occup').click(function(){
          education0008(47,2);

      });
      $$('#language0008').click(function(){
          nativeLanguage(52,2);

      });

      function nativeLanguage(row,col){
          alert("请手动再填下下拉选项以便保存");
          console.log("nativelog");
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          const nativeLanguage = document.querySelector('#nativeLanguage');
          const language = document.querySelector('#language');
          const testing_yes= document.querySelector('input[id="testing-yes"]');
          const testing_no= document.querySelector('input[id="testing-no"]');
          console.log("nativelang"+row+col+form0008[row][col]);
          nativeLanguage.selectedIndex = form0008[row][col];
          nativeLanguage.dispatchEvent(new Event('change', {bubbles:true}));
          language.selectedIndex = form0008[++row][col];
          console.log("language"+row+col+form0008[row][col]);
          language.dispatchEvent(new Event('change', {bubbles:true}));
          if (form0008[++row][col] === "Yes") {
              testing_yes.checked = true;
              testing_yes.dispatchEvent(new Event('change', {bubbles:true}));
          }else if (form0008[row][col] === "No") {
              testing_no.checked = true;
              testing_no.dispatchEvent(new Event('change', {bubbles:true}));
          }
          return row;
      }
      function education0008(row,col){
          var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          console.log("dependentInd0008 551 "+row);
          const educationLevel = document.querySelector('#educationLevel');
          const numberOfYear = document.querySelector('#numberOfYear');
          const currentOccupation = document.querySelector('#currentOccupation');
          const intendedOccupation = document.querySelector('#intendedOccupation');
          console.log("dependentInd0008 556 "+row + " " + col);
          educationLevel.selectedIndex = form0008[row][col];
          educationLevel.dispatchEvent(new Event('change', {bubbles:true}));
          imitateKeyInput(numberOfYear,"");
          imitateKeyInput(numberOfYear,form0008[++row][col]);
          imitateKeyInput(currentOccupation,form0008[++row][col]);
          imitateKeyInput(intendedOccupation,form0008[++row][col]);
          return row;

      };
      var baseDep0008 = 0;
      var cellRowDep0008 = 0;
      $$('#dep0008-1').click(async function(){
          // var form0008 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]], {header:1});
          baseDep0008 = dependent0008Indices[form0008_counter].rowIndex;
          cellRowDep0008 =baseDep0008;
          console.log("form0008[cellRowDep0008][3]"+form0008[cellRowDep0008][3]);
          if(form0008[cellRowDep0008][3] === "No"){
              var accompReason = document.querySelector('input[id="dependantDetailsForm-reasonNotAccompanying"]');
              imitateKeyInput(accompReason,form0008[cellRowDep0008][4]);
              console.log("accomreason"+form0008[cellRowDep0008][4]);
          }
          //console.log("form0008[56][3] outer"+form0008[++dependentInd0008][3]);
          if(form0008[++cellRowDep0008][3] != undefined){
              var depType = document.querySelector('select[id="dependantDetailsForm-dependantType"]');
              console.log("depType"+depType);
              depType.selectedIndex = form0008[cellRowDep0008][3];
              depType.dispatchEvent(new Event('change', {bubbles:true}));
          }

          await personalDetails0008(depPersonalDetailIndices[form0008_counter].rowIndex+1,4);

          console.log("form5406_counter",form5406_counter)
          var base = dependent0008Indices[form0008_counter].rowIndex;
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);

      });

      $$('#dep0008-2').click(async function(){
          await education0008(depEduIndices[form0008_counter].rowIndex,4);
          var base = dependent0008Indices[form0008_counter].rowIndex;
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);});

      $$('#dep0008-3').click(function(){
          console.log('#dep0008-3' + depLanguageIndices[form0008_counter].rowIndex)
          nativeLanguage(depLanguageIndices[form0008_counter].rowIndex,4);
          var base = dependent0008Indices[form0008_counter].rowIndex;
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);});

      $$('#dep0008-4').click(async function(){
          await passport(depPassportIndices[form0008_counter].rowIndex,4);
          var base = dependent0008Indices[form0008_counter].rowIndex;
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);}); //84

      $$('#dep0008-5').click(async function(){
          await nationalID(depNationalIDIndices[form0008_counter].rowIndex,4);
          console.log("#dep0008-5"+cellRowDep0008+" 4")
          var base = dependent0008Indices[form0008_counter].rowIndex;
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);
      });

      $$('#next0008').click(function (){
          ++form0008_counter;
          var base = dependent0008Indices[form0008_counter].rowIndex;
          if(form0008_counter == dependent0008Indices.length-1)
          {alert('最后一个dependent');
          }
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);

      })
      $$('#prev0008').click(function (){
          -- form0008_counter;
          var base = dependent0008Indices[form0008_counter].rowIndex;
          if(form0008_counter == 0)
          {alert('第一个dependent');
          }
          $('#texta').val('0008 person ' +form0008_counter+' '+ form0008[1+base][0]);
      })


      /*5406*/
      var child_ind = 0;
      var sibing_ind = 0;
      var trp_ind = 0;
      function findIndicesOfValue(matrix, value) {
          var indices = [];
          matrix.forEach(function (row, rowIndex) {
              row.forEach(function (cell, colIndex) {
                  if (cell === value) {
                      indices.push({ rowIndex: rowIndex, colIndex: colIndex });
                  }
              });
          });
          return indices;
      }

      $$('#nextDependant').click(function (){
          ++form5406_counter;
          var base = sectionAIndices[form5406_counter].rowIndex;
          if(form5406_counter == sectionAIndices.length-1)
          {alert('最后一个');
          }
          $('#texta').val('5406 person' +form5406_counter+' '+ form5406[1+base][0]);
          console.log("form5406_counter ",form5406_counter);
          child_ind = 0;
          sibing_ind = 0;
          trp_ind = 0;
      });

      $$('#prevDependant').click(function (){
          --form5406_counter;
          var base = sectionAIndices[form5406_counter].rowIndex;
          if(form5406_counter == 0)
          {alert('第一个');
          }
          $('#texta').val('5406 person' +form5406_counter+' '+ form5406[1+base][0]);
          console.log("form5406_counter ",form5406_counter);
          child_ind = 0;
          sibing_ind = 0;
          trp_ind = 0;
      });

      $$('#sectionA5406').click(async function(){

          console.log("5406", form5406);
          if (form5406_counter<=depNum){
              console.log(sectionAIndices);

              var base = sectionAIndices[form5406_counter].rowIndex;
              $('#texta').val('5406 person' +form5406_counter+' '+ form5406[1+base][0]);

              const country = document.querySelector('.form-input__field[name="applicantBirthplace"]');
              imitateKeyInput(country,form5406[2+sectionAIndices[form5406_counter].rowIndex][3]);
              const email = document.querySelector('.form-input__field[name="applicantEmail"]');
              imitateKeyInput(email,form5406[4+sectionAIndices[form5406_counter].rowIndex][3]);
              const address = document.querySelector('.form-textarea__field[name="applicantAddress"]');
              imitateKeyInput(address,form5406[5+sectionAIndices[form5406_counter].rowIndex][3]);

              var martial = document.querySelector('select[id="applicantMaritalStatus"]');
              console.log("martial", martial);

              martial.selectedIndex = form5406[3+sectionAIndices[form5406_counter].rowIndex][3];
              martial.dispatchEvent(changeEvent);

              // Spouse
              const partner_fullname = document.querySelector('.form-input__field[name="partnerFullName"]');
              imitateKeyInput(partner_fullname,form5406[7+sectionAIndices[form5406_counter].rowIndex][3]);
              await delay(500);
              const partnerDOB = document.querySelector('.form-datepicker__field[name="partnerDOB"]');
              imitateKeyInput(partnerDOB,form5406[8+sectionAIndices[form5406_counter].rowIndex][3]);

              const p_country = document.querySelector('.form-input__field[name="partnerBirthplace"]');
              imitateKeyInput(p_country,form5406[9+sectionAIndices[form5406_counter].rowIndex][3]);
              const p_email = document.querySelector('.form-input__field[name="partnerEmail"]');
              imitateKeyInput(p_email,form5406[11+sectionAIndices[form5406_counter].rowIndex][3]);
              const p_address = document.querySelector('.form-textarea__field[name="partnerAddress"]');
              imitateKeyInput(p_address,form5406[12+sectionAIndices[form5406_counter].rowIndex][3]);
              var p_martial = document.querySelector('select[id="partnerMaritalStatus"]');
              p_martial.selectedIndex = form5406[10+sectionAIndices[form5406_counter].rowIndex][3];
              p_martial.dispatchEvent(changeEvent);

              // Mother
              const mother_fullname = document.querySelector('.form-input__field[name="motherFullName"]');
              imitateKeyInput(mother_fullname,form5406[14+sectionAIndices[form5406_counter].rowIndex][3]);
              await delay(500);
              const motherDOB = document.querySelector('.form-datepicker__field[name="motherDOB"]');
              imitateKeyInput(motherDOB,form5406[15+sectionAIndices[form5406_counter].rowIndex][3]);

              const m_country = document.querySelector('.form-input__field[name="motherBirthplace"]');
              imitateKeyInput(m_country,form5406[16+sectionAIndices[form5406_counter].rowIndex][3]);
              const m_email = document.querySelector('.form-input__field[name="motherEmail"]');
              imitateKeyInput(m_email,form5406[18+sectionAIndices[form5406_counter].rowIndex][3]);
              const m_address = document.querySelector('.form-textarea__field[name="motherAddress"]');
              imitateKeyInput(m_address,form5406[19+sectionAIndices[form5406_counter].rowIndex][3]);
              var m_martial = document.querySelector('select[id="motherMaritalStatus"]');
              m_martial.selectedIndex = form5406[17+sectionAIndices[form5406_counter].rowIndex][3];
              m_martial.dispatchEvent(changeEvent);

              // Father
              const father_fullname = document.querySelector('.form-input__field[name="fatherFullName"]');
              imitateKeyInput(father_fullname,form5406[21+sectionAIndices[form5406_counter].rowIndex][3]);
              await delay(500);
              const fatherDOB = document.querySelector('.form-datepicker__field[name="fatherDOB"]');
              imitateKeyInput(fatherDOB,form5406[22+sectionAIndices[form5406_counter].rowIndex][3]);

              const f_country = document.querySelector('.form-input__field[name="fatherBirthplace"]');
              imitateKeyInput(f_country,form5406[23+sectionAIndices[form5406_counter].rowIndex][3]);
              const f_email = document.querySelector('.form-input__field[name="fatherEmail"]');
              imitateKeyInput(f_email,form5406[25+sectionAIndices[form5406_counter].rowIndex][3]);
              const f_address = document.querySelector('.form-textarea__field[name="fatherAddress"]');
              imitateKeyInput(f_address,form5406[26+sectionAIndices[form5406_counter].rowIndex][3]);
              var f_martial = document.querySelector('select[id="fatherMaritalStatus"]');
              f_martial.selectedIndex = form5406[24+sectionAIndices[form5406_counter].rowIndex][3];
              f_martial.dispatchEvent(changeEvent);
          }
      })

      $$('#sectionB5406').click(async function(){
          console.log("index:", sectionAIndices[form5406_counter].rowIndex);
          var form5406 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[3]], {header:1});
          const c_no = form5406[27+sectionAIndices[form5406_counter].rowIndex][3]
          //             var child_DOB;
          //             var relationship;
          //             var child_fullname;
          //             var c_martial;
          //             var c_country;
          //             var c_email;
          //             var c_address;
          if (child_ind<c_no && form5406_counter<=depNum){
              var base = sectionAIndices[form5406_counter].rowIndex;
              $('#texta').val('5406 person' +form5406_counter+' '+ form5406[1+base][0]);
              var relationship = document.querySelector('.form-input__field[name="relationship'+child_ind+'"]');
              imitateKeyInput(relationship,form5406[29+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);
              var child_fullname = document.querySelector('.form-input__field[name="fullName'+child_ind+'"]');
              imitateKeyInput(child_fullname,form5406[30+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);

              await delay(500);
              var child_DOB = document.querySelector('input[name="dob'+child_ind+'"]');
              console.log("child_DOB", 'input[name="dob'+child_ind+'"]');
              imitateKeyInput(child_DOB,form5406[31+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);

              var c_country = document.querySelector('.form-input__field[name="countryOfBirth'+child_ind+'"]');
              imitateKeyInput(c_country,form5406[32+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);
              var c_email = document.querySelector('.form-input__field[name="emailAddress'+child_ind+'"]');
              imitateKeyInput(c_email,form5406[34+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);
              var c_address = document.querySelector('.form-textarea__field[name="address'+child_ind+'"]');
              imitateKeyInput(c_address,form5406[35+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3]);
              var c_martial = document.querySelector('select[id="maritalStatus'+child_ind+'"]');
              c_martial.selectedIndex = form5406[33+child_ind*8+sectionAIndices[form5406_counter].rowIndex][3];
              c_martial.dispatchEvent(changeEvent);
              await delay(500);
              child_ind = child_ind + 1;
              var child_left = c_no - child_ind;
              if(child_left >0){
              alert('记得点Add another,还有'+child_left+'次');
          } else if(child_left ===0){
              alert('结束，按Save and continue');
          }
          }
      })


      $$('#sectionC5406').click(async function(){
          var form5406 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[3]], {header:1});
          const s_no = form5406[sectionCIndices[form5406_counter].rowIndex][3];
          var sibing_DOB;
          var relationship;
          var sibing_fullname;
          var s_martial;
          var s_country;
          var s_email;
          var s_address;
          if (sibing_ind<s_no && form5406_counter<=depNum){
              var base = sectionAIndices[form5406_counter].rowIndex;
              $('#texta').val('5406 person' +form5406_counter+' '+ form5406[1+base][0]);

              relationship = document.querySelector('.form-input__field[name="relationship'+sibing_ind+'"]');
              console.log("sectionCIndices[form5406_counter]", sectionCIndices[form5406_counter]);
              imitateKeyInput(relationship,form5406[sectionCIndices[form5406_counter].rowIndex+2+sibing_ind*8][3]);
              sibing_fullname = document.querySelector('.form-input__field[name="fullName'+sibing_ind+'"]');
              imitateKeyInput(sibing_fullname,form5406[sectionCIndices[form5406_counter].rowIndex+3+sibing_ind*8][3]);

              await delay(500);
              sibing_DOB = document.querySelector('input[name="dob'+sibing_ind+'"]');
              console.log("sibing_DOB", form5406[sectionCIndices[form5406_counter].rowIndex+4+sibing_ind*8][3]);
              imitateKeyInput(sibing_DOB,form5406[sectionCIndices[form5406_counter].rowIndex+4+sibing_ind*8][3]);

              s_country = document.querySelector('.form-input__field[name="countryOfBirth'+sibing_ind+'"]');
              imitateKeyInput(s_country,form5406[sectionCIndices[form5406_counter].rowIndex+5+sibing_ind*8][3]);
              s_email = document.querySelector('.form-input__field[name="emailAddress'+sibing_ind+'"]');
              imitateKeyInput(s_email,form5406[sectionCIndices[form5406_counter].rowIndex+7+sibing_ind*8][3]);
              s_address = document.querySelector('.form-textarea__field[name="address'+sibing_ind+'"]');
              imitateKeyInput(s_address,form5406[sectionCIndices[form5406_counter].rowIndex+8+sibing_ind*8][3]);
              s_martial = document.querySelector('select[id="maritalStatus'+sibing_ind+'"]');
              s_martial.selectedIndex = form5406[sectionCIndices[form5406_counter].rowIndex+6+sibing_ind*8][3];
              s_martial.dispatchEvent(changeEvent);
              await delay(500);
              sibing_ind = sibing_ind + 1;
              var sibiling_left = s_no - sibing_ind;
              if(sibiling_left >0){
              alert('记得点Add another,还有'+sibiling_left+'次');
          } else if(sibiling_left ===0){
              alert('结束，按Save and continue');
          }

          }
      })


      var name_counter = 0;
      $$('#sectionA5562').click(async function(){
          var n_trip = form5562[4+sectionAIndices5562[form5562_counter].rowIndex][4];
          var data = form5562[4+sectionAIndices5562[form5562_counter].rowIndex][2];
          console.log("5562", form5562);
          console.log("n_trip", n_trip);
          console.log("trp_ind", trp_ind);
          console.log("depNum", depNum);
          console.log("data", data);
          $('#texta').val('5562 A');

          if (name_counter<1){
                  const familyName = document.querySelector('.form-input__field[name="familyName"]');
                  console.log("row index:", sectionAIndices5562[form5562_counter]);
                  imitateKeyInput(familyName,form5562[2+sectionAIndices5562[form5562_counter].rowIndex][3]);
                  const givenName = document.querySelector('.form-input__field[name="givenName"]');
                  imitateKeyInput(givenName,form5562[3+sectionAIndices5562[form5562_counter].rowIndex][3]);

              }

          const yesInput = document.querySelector('.checkbox__input[name="haveNotTravelled"]');
          console.log("yesInput", yesInput);
          if (data === "Yes") {
                      // Select "yes" option

                      yesInput.checked = true;
                      yesInput.dispatchEvent(new Event('change', {bubbles:true}));
                  }

          await delay(500);
          if (form5562_counter<=depNum && trp_ind<n_trip){
              console.log(sectionAIndices);


              await delay(500);
              var from_date = document.querySelector('input[name="trip-from-'+trp_ind+'"]');
              console.log("from_date", 'input[name="trip-from-'+trp_ind+'"]');
              imitateKeyInput(from_date,form5562[sectionAIndices5562[form5562_counter].rowIndex+6+trp_ind*6][3]);

              await delay(500);
              var to_date = document.querySelector('input[name="trip-to-'+trp_ind+'"]');
              console.log("to_date", 'input[name="trip-to-'+trp_ind+'"]');
              imitateKeyInput(to_date,form5562[sectionAIndices5562[form5562_counter].rowIndex+7+trp_ind*6][3]);

              var destination = document.querySelector('input[name="trip-destination-'+trp_ind+'"]');
              console.log("destination", 'input[name="trip-destination-'+trp_ind+'"]');
              imitateKeyInput(destination,form5562[sectionAIndices5562[form5562_counter].rowIndex+8+trp_ind*6][3]);

              var length = document.querySelector('input[name="trip-length-'+trp_ind+'"]');
              console.log("length", 'input[name="trip-length-'+trp_ind+'"]');
              imitateKeyInput(length,form5562[sectionAIndices5562[form5562_counter].rowIndex+9+trp_ind*6][3]);

              var purpose = document.querySelector('input[name="trip-purpose-'+trp_ind+'"]');
              console.log("purpose", 'input[name="trip-purpose-'+trp_ind+'"]');
              imitateKeyInput(purpose,form5562[sectionAIndices5562[form5562_counter].rowIndex+10+trp_ind*6][3]);

              await delay(500);
              trp_ind = trp_ind + 1;
              name_counter++;
              var trp_left = trp_left = n_trip - trp_ind;;
              if(trp_left >0){
              alert('记得点Add another,还有'+trp_left+'次');
          } else if(trp_left ===0){
              alert('结束，按Save and continue');
          }
          }
      })
      var trpB_ind = 0;
      $$('#sectionB5562').click(async function(){
          var data = form5562[1+sectionBIndices5562[form5562_counter].rowIndex][2];
          var n_trip = form5562[1+sectionBIndices5562[form5562_counter].rowIndex][4];
          console.log("5562", form5562);
          $('#texta').val('5562 B');
          await delay(200);
          const yesInput = document.querySelector('.checkbox__input[name="haveNotTravelled"]');
          console.log("5562", form5562);
          if (data === "Yes") {
              // Select "yes" option
              yesInput.checked = true;
              yesInput.dispatchEvent(new Event('change', {bubbles:true}));
          }

          await delay(500);
          if (form5562_counter<=depNum && trpB_ind<n_trip){

              var from_date = document.querySelector('input[name="trip-from-'+trpB_ind+'"]');
              console.log("from_date", 'input[name="trip-from-'+trpB_ind+'"]');
              imitateKeyInput(from_date,form5562[sectionBIndices5562[form5562_counter].rowIndex+3+trpB_ind*6][3]);

              await delay(500);
              var to_date = document.querySelector('input[name="trip-to-'+trpB_ind+'"]');
              console.log("to_date", 'input[name="trip-to-'+trpB_ind+'"]');
              imitateKeyInput(to_date,form5562[sectionBIndices5562[form5562_counter].rowIndex+4+trpB_ind*6][3]);

              var destination = document.querySelector('input[name="sectionB-destination-'+trpB_ind+'"]');
              console.log("destination", 'input[name="trip-destination-'+trpB_ind+'"]');
              imitateKeyInput(destination,form5562[sectionBIndices5562[form5562_counter].rowIndex+5+trpB_ind*6][3]);

              var length = document.querySelector('input[name="trip-length-'+trpB_ind+'"]');
              console.log("length", 'input[name="trip-length-'+trpB_ind+'"]');
              imitateKeyInput(length,form5562[sectionBIndices5562[form5562_counter].rowIndex+6+trpB_ind*6][3]);

              var purpose = document.querySelector('input[name="trip-purpose-'+trpB_ind+'"]');
              console.log("purpose", 'input[name="trip-purpose-'+trpB_ind+'"]');
              imitateKeyInput(purpose,form5562[sectionBIndices5562[form5562_counter].rowIndex+7+trpB_ind*6][3]);

              await delay(500);
              trpB_ind = trpB_ind + 1;
              var trpB_left = n_trip - trpB_ind;;
              if(trpB_left >0){
              alert('记得点Add another,还有'+trpB_left+'次');
          } else if(trpB_left ===0){
              alert('结束，按Save and continue');
          }
          }
      })

      var trpC_ind = 0;
      $$('#sectionC5562').click(async function(){
          var data = form5562[1+sectionCIndices5562[form5562_counter].rowIndex][2];
          var n_trip = form5562[1+sectionCIndices5562[form5562_counter].rowIndex][4];
          console.log("5562", form5562);
          $('#texta').val('5562 C');
          await delay(200);
          const yesInput = document.querySelector('.checkbox__input[name="haveNotTravelled"]');
          console.log("5562", form5562);
          if (data === "Yes") {
              // Select "yes" option
              yesInput.checked = true;
              yesInput.dispatchEvent(new Event('change', {bubbles:true}));
          }

          await delay(500);
          if (form5562_counter<=depNum && trpC_ind<n_trip){

              const givenName = document.querySelector('input[name="trip-givenName-'+trpC_ind+'"]');
              imitateKeyInput(givenName,form5562[3+sectionCIndices5562[form5562_counter].rowIndex][3]);

              await delay(500);
              var from_date = document.querySelector('input[name="trip-from-'+trpC_ind+'"]');
              console.log("from_date", 'input[name="trip-from-'+trpC_ind+'"]');
              imitateKeyInput(from_date,form5562[sectionCIndices5562[form5562_counter].rowIndex+4+trpC_ind*7][3]);

              await delay(500);
              var to_date = document.querySelector('input[name="trip-to-'+trpC_ind+'"]');
              console.log("to_date", 'input[name="trip-to-'+trpC_ind+'"]');
              imitateKeyInput(to_date,form5562[sectionCIndices5562[form5562_counter].rowIndex+5+trpC_ind*7][3]);

              var destination = document.querySelector('input[name="sectionC-destination-'+trpC_ind+'"]');
              console.log("destination", 'input[name="sectionC-destination-'+trpC_ind+'"]');
              imitateKeyInput(destination,form5562[sectionCIndices5562[form5562_counter].rowIndex+6+trpC_ind*7][3]);

              var length = document.querySelector('input[name="trip-length-'+trpC_ind+'"]');
              console.log("length", 'input[name="trip-length-'+trpC_ind+'"]');
              imitateKeyInput(length,form5562[sectionCIndices5562[form5562_counter].rowIndex+7+trpC_ind*7][3]);

              var purpose = document.querySelector('input[name="trip-purpose-'+trpC_ind+'"]');
              console.log("purpose", 'input[name="trip-purpose-'+trpC_ind+'"]');
              imitateKeyInput(purpose,form5562[sectionCIndices5562[form5562_counter].rowIndex+8+trpC_ind*7][3]);

              await delay(500);
              trpC_ind = trpC_ind + 1;
              var trpC_left = n_trip - trpC_ind;;
              if(trpC_left >0){
              alert('记得点Add another,还有'+trpC_left+'次');
          } else if(trpC_left ===0){
              alert('结束，按Save and continue');
          }
          }
      })


      /*5669*/
      $$("#persDetails5669").click(function(){
          console.log("persDetails5669");
          personalDetails5669();
      })
      function personalDetails5669(){
          var base = personalDetailIndices[form5569_counter].rowIndex;

          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);

          var nativeFullName= document.querySelector('#nativeFullName');
          imitateKeyInput(nativeFullName, form5669[2+base][4]);
          //father info
          var familyNameFather= document.querySelector('#familyNameFather');
          imitateKeyInput(familyNameFather, form5669[4+base][4]);
          var givenNameFather= document.querySelector('#givenNameFather');
          imitateKeyInput(givenNameFather, form5669[5+base][4]);
          var sectionAFormFatherDOB= document.querySelector('#sectionAFormFatherDOB');
          imitateKeyInput(sectionAFormFatherDOB, form5669[6+base][4]);
          var sectionAFormFatherDeceasedDate= document.querySelector('#sectionAFormFatherDeceasedDate');
          imitateKeyInput(sectionAFormFatherDeceasedDate, form5669[7+base][4]);
          var sectionAFormFatherCityOfBirth= document.querySelector('#sectionAFormFatherCityOfBirth');
          imitateKeyInput(sectionAFormFatherCityOfBirth, form5669[8+base][4]);
          var sectionAFormFatherCountryOfBirth= document.querySelector('#sectionAFormFatherCountryOfBirth');
          imitateKeyInput(sectionAFormFatherCountryOfBirth, form5669[9+base][4]);
          // mother info
          var familyNameMother= document.querySelector('#familyNameMother');
          imitateKeyInput(familyNameMother, form5669[11+base][4]);
          var givenNameMother= document.querySelector('#givenNameMother');
          imitateKeyInput(givenNameMother, form5669[12+base][4]);
          var sectionAFormMotherDOB= document.querySelector('#sectionAFormMotherDOB');
          imitateKeyInput(sectionAFormMotherDOB, form5669[13+base][4]);
          var sectionAFormMotherDeceasedDate= document.querySelector('#sectionAFormMotherDeceasedDate');
          imitateKeyInput(sectionAFormMotherDeceasedDate, form5669[14+base][4]);
          var sectionAFormMotherCityOfBirth= document.querySelector('#sectionAFormMotherCityOfBirth');
          imitateKeyInput(sectionAFormMotherCityOfBirth, form5669[15+base][4]);
          var sectionAFormMotherCountryOfBirth= document.querySelector('#sectionAFormMotherCountryOfBirth');
          imitateKeyInput(sectionAFormMotherCountryOfBirth, form5669[16+base][4]);
      };
      $$("#Quesitonaire5669").click(function(){
          console.log("Quesitonaire");
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          quesitonaire(4);
      })
      function ifHelper(yes,no,val){
          if(val === "Yes"){
              yes.checked = true;
              yes.dispatchEvent(new Event('change', {bubbles:true}));
          } else if(val === "No"){
              no.checked = true;
              no.dispatchEvent(new Event('change', {bubbles:true}));
          }
      }
      function quesitonaire(col){
          var base = QuestionnaireIndices[form5569_counter].rowIndex;

          var isConvictedInCanada_yes= document.querySelector('#isConvictedInCanada_yes');
          var isConvictedInCanada_no= document.querySelector('#isConvictedInCanada_no');
          var isConvictedInCanada = form5669[1+base][col];
          ifHelper(isConvictedInCanada_yes,isConvictedInCanada_no,isConvictedInCanada);

          var isConvictedOutsideCanada_yes= document.querySelector('#isConvictedOutsideCanada_yes');
          var isConvictedOutsideCanada_no= document.querySelector('#isConvictedOutsideCanada_no');
          var isConvictedOutsideCanada = form5669[2+base][col];
          ifHelper(isConvictedOutsideCanada_yes,isConvictedOutsideCanada_no,isConvictedOutsideCanada);

          var isClaimedRefugeeProtection_yes= document.querySelector('#isClaimedRefugeeProtection_yes');
          var isClaimedRefugeeProtection_no= document.querySelector('#isClaimedRefugeeProtection_no');
          var isClaimedRefugeeProtection = form5669[3+base][col];
          ifHelper(isClaimedRefugeeProtection_yes,isClaimedRefugeeProtection_no,isClaimedRefugeeProtection);

          var isRefusedRefugeeOrVisa_yes= document.querySelector('#isRefusedRefugeeOrVisa_yes');
          var isRefusedRefugeeOrVisa_no= document.querySelector('#isRefusedRefugeeOrVisa_no');
          var isRefusedRefugeeOrVisa = form5669[4+base][col];
          ifHelper(isRefusedRefugeeOrVisa_yes,isRefusedRefugeeOrVisa_no,isRefusedRefugeeOrVisa);

          var isOrderedToLeaveCountry_yes= document.querySelector('#isOrderedToLeaveCountry_yes');
          var isOrderedToLeaveCountry_no= document.querySelector('#isOrderedToLeaveCountry_no');
          var isOrderedToLeaveCountry = form5669[5+base][col];
          ifHelper(isOrderedToLeaveCountry_yes,isOrderedToLeaveCountry_no,isOrderedToLeaveCountry);

          var isWarCriminal_yes= document.querySelector('#isWarCriminal_yes');
          var isWarCriminal_no= document.querySelector('#isWarCriminal_no');
          var isWarCriminal = form5669[6+base][col];
          ifHelper(isWarCriminal_yes,isWarCriminal_no,isWarCriminal);

          var isCommittedActOfViolence_yes= document.querySelector('#isCommittedActOfViolence_yes');
          var isCommittedActOfViolence_no= document.querySelector('#isCommittedActOfViolence_no');
          var isCommittedActOfViolence = form5669[7+base][col];
          ifHelper(isCommittedActOfViolence_yes,isCommittedActOfViolence_no,isCommittedActOfViolence);


          var isAssociatedWithViolentGroup_yes= document.querySelector('#isAssociatedWithViolentGroup_yes');
          var isAssociatedWithViolentGroup_no= document.querySelector('#isAssociatedWithViolentGroup_no');
          var isAssociatedWithViolentGroup = form5669[8+base][col];
          ifHelper(isAssociatedWithViolentGroup_yes,isAssociatedWithViolentGroup_no,isAssociatedWithViolentGroup);

          var isMemberOfCriminalOrg_yes= document.querySelector('#isMemberOfCriminalOrg_yes');
          var isMemberOfCriminalOrg_no= document.querySelector('#isMemberOfCriminalOrg_no');
          var isMemberOfCriminalOrg = form5669[9+base][col];
          ifHelper(isMemberOfCriminalOrg_yes,isMemberOfCriminalOrg_no,isMemberOfCriminalOrg);

          var isDetainedOrJailed_yes= document.querySelector('#isDetainedOrJailed_yes');
          var isDetainedOrJailed_no= document.querySelector('#isDetainedOrJailed_no');
          var isDetainedOrJailed = form5669[10+base][col];
          ifHelper(isDetainedOrJailed_yes,isDetainedOrJailed_no,isDetainedOrJailed );

          var isPhysicalOrMentalDisorder_yes= document.querySelector('#isPhysicalOrMentalDisorder_yes');
          var isPhysicalOrMentalDisorder_no= document.querySelector('#isPhysicalOrMentalDisorder_no');
          var isPhysicalOrMentalDisorder = form5669[11+base][col];
          ifHelper(isPhysicalOrMentalDisorder_yes,isPhysicalOrMentalDisorder_no,isPhysicalOrMentalDisorder );

          var additionalDetails= document.querySelector('#additionalDetails');
          imitateKeyInput(additionalDetails, form5669[12+base][col]);
      }
      $$("#Education5669").click(function(){
          console.log("Education5669");
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          education();
      })

      var eduIteTime = 0;
      var eduBase5669 =0;

      function education(){
          eduIteTime = 0
          eduBase5669 = educationIndices[form5569_counter].rowIndex;
          var base = eduBase5669;
          var elementarySchoolYears= document.querySelector('#elementarySchoolYears');
          imitateKeyInput(elementarySchoolYears, form5669[2+base][4]);
          var secondarySchoolYears= document.querySelector('#secondarySchoolYears');
          imitateKeyInput(secondarySchoolYears, form5669[3+base][4]);
          var universityAndCollegeYears= document.querySelector('#universityAndCollegeYears');
          imitateKeyInput(universityAndCollegeYears, form5669[4+base][4]);
          var otherSchoolYears= document.querySelector('#otherSchoolYears');
          imitateKeyInput(otherSchoolYears, form5669[5+base][4]);
      }

      $$("#Edu-add5669").click(function(){
          console.log("Education5669-add");
          edu_add();

      })
      var eduLeft = 0;
      function edu_add(){
          eduBase5669 = educationIndices[form5569_counter].rowIndex;
          var base = eduBase5669;
          var eduIndex = form5669[base][4];
          if(eduIteTime >=eduIndex) return;

          var from0= document.querySelector('#from'+eduIteTime);
          console.log("eduIteTime 302 "+eduIteTime);
          imitateKeyInput(from0, form5669[7+base+eduIteTime*6][5]);
          var to0= document.querySelector('#to'+eduIteTime);
          imitateKeyInput(to0, form5669[8+base+eduIteTime*6][5]);
          var nameOfInstitution0= document.querySelector('#nameOfInstitution'+eduIteTime);
          imitateKeyInput(nameOfInstitution0, form5669[9+base+eduIteTime*6][5]);
          var cityAndCountry0= document.querySelector('#cityAndCountry'+eduIteTime);
          imitateKeyInput(cityAndCountry0, form5669[10+base+eduIteTime*6][5]);
          var typeOfDiploma0= document.querySelector('#typeOfDiploma'+eduIteTime);
          imitateKeyInput(typeOfDiploma0, form5669[11+base+eduIteTime*6][5]);
          var fieldOfStudy0= document.querySelector('#fieldOfStudy'+eduIteTime);
          imitateKeyInput(fieldOfStudy0, form5669[12+base+eduIteTime*6][5]);
          eduIteTime++;
          eduLeft =eduIndex - eduIteTime;
          if(eduLeft >0){
              alert('记得点Add another,还有'+eduLeft+'次');
          } else if(eduLeft ===0){
              alert('结束，按Save and continue');
          }
      }

      var persHistIteTime = 0;
      var persHistBase5669 = 0;
      var perHistLeft = 0;
      $$("#persHist5669").click(function(){
          persHistBase5669 = persHistIndices[form5569_counter].rowIndex;
          console.log("persHist5669");
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          persHist5669();
      })

      function persHist5669(){
          var persHistInd = form5669[1+persHistBase5669][3];
          if(persHistIteTime >= persHistInd) return;
          var from0= document.querySelector('#from'+persHistIteTime);
          imitateKeyInput(from0, form5669[2+persHistBase5669+persHistIteTime*6][4]);
          var to0= document.querySelector('#to'+persHistIteTime);
          imitateKeyInput(to0, form5669[3+persHistBase5669+persHistIteTime*6][4]);
          var activity0= document.querySelector('#activity'+persHistIteTime);
          imitateKeyInput(activity0, form5669[4+persHistBase5669+persHistIteTime*6][4]);
          var cityAndCountry0= document.querySelector('#cityAndCountry'+persHistIteTime);
          imitateKeyInput(cityAndCountry0, form5669[5+persHistBase5669+persHistIteTime*6][4]);
          var status0= document.querySelector('#status'+persHistIteTime);
          imitateKeyInput(status0, form5669[6+persHistBase5669+persHistIteTime*6][4]);
          var nameOfEmployerOrSchool0= document.querySelector('#nameOfEmployerOrSchool'+persHistIteTime);
          imitateKeyInput(nameOfEmployerOrSchool0, form5669[7+persHistBase5669+persHistIteTime*6][4]);
          persHistIteTime++;
          perHistLeft = persHistInd - persHistIteTime
          console.log("persHistInd"+persHistInd);
          console.log("persHistIteTime"+persHistIteTime);
          console.log("perHistLeft"+perHistLeft);
          if(perHistLeft >0){
              alert('记得点Add another,还有'+perHistLeft+'次');
          } else {
              alert('结束，按Save and continue');
          }
      }

      var membershipIteTime = 0;
      var membershipbase5669 = 0;
      $$("#membership5669").click(function(){
          membershipbase5669 = membershipIndices[form5569_counter].rowIndex;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          membership5669();
      })
      function membership5669(){
          console.log("membershipIteTime >= form5669[1+membershipbase5669][3]:"+membershipIteTime +" " +form5669[1+membershipbase5669][3]);

          if(membershipIteTime >= form5669[1+membershipbase5669][3]) return;
          var from0= document.querySelector('#from'+membershipIteTime);
          imitateKeyInput(from0, form5669[2+membershipbase5669+membershipIteTime*6][4]);
          var to0= document.querySelector('#to'+membershipIteTime);
          imitateKeyInput(to0, form5669[3+membershipbase5669+membershipIteTime*6][4]);
          var nameOfOrganization0= document.querySelector('#nameOfOrganization'+membershipIteTime);
          imitateKeyInput(nameOfOrganization0, form5669[4+membershipbase5669+membershipIteTime*6][4]);
          var typeOfOrganization0= document.querySelector('#typeOfOrganization'+membershipIteTime);
          imitateKeyInput(typeOfOrganization0, form5669[5+membershipbase5669+membershipIteTime*6][4]);
          var activities0= document.querySelector('#activities'+membershipIteTime);
          imitateKeyInput(activities0, form5669[6+membershipbase5669+membershipIteTime*6][4]);
          var cityAndCountry0= document.querySelector('#cityAndCountry'+membershipIteTime);
          imitateKeyInput(cityAndCountry0, form5669[7+membershipbase5669+membershipIteTime*6][4]);
          membershipIteTime++;
      }

      var govPositionbase5669 = 0;
      var govPositionIteTime = 0;
      $$("#govPosition5669").click(function(){
          govPositionbase5669 = govermentIndices[form5569_counter].rowIndex;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          govPosition5669();
      })
      function govPosition5669(){
          if(govPositionIteTime>= form5669[1+govPositionbase5669][3]) return;
          var dateFrom0= document.querySelector('#dateFrom'+govPositionIteTime);
          imitateKeyInput(dateFrom0, form5669[2+govPositionbase5669+govPositionIteTime*5][4]);
          var to0= document.querySelector('#to'+govPositionIteTime);
          imitateKeyInput(to0, form5669[3+govPositionbase5669+govPositionIteTime*5][4]);
          var cityAndCountry0= document.querySelector('#cityAndCountry'+govPositionIteTime);
          imitateKeyInput(cityAndCountry0, form5669[4+govPositionbase5669+govPositionIteTime*5][4]);
          var department0= document.querySelector('#department'+govPositionIteTime);
          imitateKeyInput(department0, form5669[5+govPositionbase5669+govPositionIteTime*5][4]);
          var activities0= document.querySelector('#activities'+govPositionIteTime);
          imitateKeyInput(activities0, form5669[6+govPositionbase5669+govPositionIteTime*5][4]);
          govPositionIteTime++;
      }

      var militarybase5669 = 0;
      var militaryIteTime = 0;
      $$("#military5669").click(function(){
          militarybase5669 = militaryIndices[form5569_counter].rowIndex;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          military5669();
      })
      function military5669(){
          if(militaryIteTime>= form5669[1+militarybase5669][3]) return;
          var country0= document.querySelector('#country'+militaryIteTime);
          imitateKeyInput(country0, form5669[2+militarybase5669+militaryIteTime*7][4]);
          var branchOfService0= document.querySelector('#branchOfService'+militaryIteTime);
          imitateKeyInput(branchOfService0, form5669[3+militarybase5669+militaryIteTime*7][4]);
          var from0= document.querySelector('#from'+militaryIteTime);
          imitateKeyInput(from0, form5669[4+militarybase5669+militaryIteTime*7][4]);
          var to0= document.querySelector('#to'+militaryIteTime);
          imitateKeyInput(to0, form5669[5+militarybase5669+militaryIteTime*7][4]);
          var rank0= document.querySelector('#rank'+militaryIteTime);
          imitateKeyInput(rank0, form5669[6+militarybase5669+militaryIteTime*7][4]);
          var reasonsEndService0= document.querySelector('#reasonsEndService'+militaryIteTime);
          imitateKeyInput(reasonsEndService0, form5669[7+militarybase5669+militaryIteTime*7][4]);
          var combatDetails0= document.querySelector('#combatDetails'+militaryIteTime);
          imitateKeyInput(combatDetails0, form5669[8+militarybase5669+militaryIteTime*7][4]);

          militaryIteTime++;
      }

      var addressbase5669 = 0;
      var addressIteTime = 0;
      var addressLeft = 0;
      $$("#address5669").click(function(){
          addressbase5669 = addresseIndices[form5569_counter].rowIndex;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          address5669();
      })
      function address5669(){
          var addressInd = form5669[1+addressbase5669][3];
          if(addressIteTime>= addressInd) return;
          var from0= document.querySelector('#from'+addressIteTime);
          imitateKeyInput(from0, form5669[2+addressbase5669+addressIteTime*7][4]);
          var to0= document.querySelector('#to'+addressIteTime);
          imitateKeyInput(to0, form5669[3+addressbase5669+addressIteTime*7][4]);
          var street= document.querySelector('#street'+addressIteTime);
          imitateKeyInput(street, form5669[4+addressbase5669+addressIteTime*7][4]);
          var city= document.querySelector('#city'+addressIteTime);
          imitateKeyInput(city, form5669[5+addressbase5669+addressIteTime*7][4]);
          var provinceOrState= document.querySelector('#provinceOrState'+addressIteTime);
          imitateKeyInput(provinceOrState, form5669[6+addressbase5669+addressIteTime*7][4]);
          var country= document.querySelector('#country'+addressIteTime);
          imitateKeyInput(country, form5669[7+addressbase5669+addressIteTime*7][4]);
          var postalCode= document.querySelector('#postalCode'+addressIteTime);
          imitateKeyInput(postalCode, form5669[8+addressbase5669+addressIteTime*7][4]);
          addressIteTime++;
          addressLeft = addressInd - addressIteTime;
          if(addressLeft >0){
              alert('记得点Add another,还有'+addressLeft+'次');
          } else {
              alert('结束，按Save and continue');
          }

      }
      $$("#next5669").click(function(){
          ++form5569_counter;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          if(form5569_counter == personalDetailIndices.length-1)
          {alert('最后一个');
          }
          $('#texta').val('5669 person ' +form5569_counter+' '+ form5669[2+base][4]);
          addressIteTime=0;
          militaryIteTime=0;
          govPositionIteTime=0;
          membershipIteTime=0;
          persHistIteTime=0;
          eduIteTime=0;
      })
      $$("#prev5669").click(function(){
          --form5569_counter;
          var base = personalDetailIndices[form5569_counter].rowIndex;
          if(form5569_counter == 0)
          {alert('第一个');
          }
          $('#texta').val('5669 person ' + form5569_counter +' '+ form5669[2+base][4]);
          addressIteTime=0;
          militaryIteTime=0;
          govPositionIteTime=0;
          membershipIteTime=0;
          persHistIteTime=0;
          eduIteTime=0;
      })
  })
})();





