/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    // document.getElementById("insert-paragraph").onclick = insertParagraph;
    // document.getElementById("apply-style").onclick = applyStyle;
    // document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    // document.getElementById("change-font").onclick = changeFont;
    // document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    // document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    // document.getElementById("replace-text").onclick = replaceText;

    document.getElementById('para').addEventListener('change', function() {
      console.log('You selected: ', this.value);
      insertParagraph(this.value);
    });

    let select = document.getElementById('list');
    //let options = readtextFile().title;

    readtextFile().then(data => {
      //console.log(data[0].title);
      for (let i=0; i<data.length; i++) {  
        let opt = data[i].title;
        var el = document.createElement("option");
        el.textContent = opt;
        el.value = opt;
        select.appendChild(el);
      
  } 
    });

    document.getElementById('list').addEventListener('change', function() {
      console.log('You selected: ', this.value);
      insertParagraph2(this.value);
    });



  //   for (let i=0; i<options.length; i++) {  
  //       let opt = options[i];
  //       var el = document.createElement("option");
  //       el.textContent = opt;
  //       el.value = opt;
  //       select.appendChild(el);
      
  // }

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph2(selection) {
  Word.run(function (context) {
    const docBody = context.document.body;
    docBody.insertParagraph(selection,"End");
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}


async function insertParagraph(selection) {
  Word.run(function (context) {

    const docBody = context.document.body;

    let p1 = "This is paragraph 1";
    let p2 = "This is paragraph 2";
    let p3 = "This is paragraph 3";
    let p4 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur vestibulum est ut sagittis sagittis. Etiam a neque malesuada, tincidunt sem at, feugiat felis. Maecenas vel tortor lectus. Donec vel tellus mattis, aliquet nisi eget, pharetra libero. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Mauris ultrices lacinia dui id posuere. Aenean mattis dui quis condimentum interdum. Donec nec elit eu metus tempus euismod. Sed rhoncus posuere est, a varius odio. Quisque ut posuere nisi. Vivamus iaculis iaculis condimentum. Vivamus nibh libero, feugiat non nisi sed, accumsan dictum quam.";

    if (selection === "p1")
    {
      //let url = 'https://jsonplaceholder.typicode.com/posts';

      
      readtextFile().then(data => {
        //docBody.insertParagraph(data[0].title,"End")
        console.log(data[0].title);
        Office.context.document.setSelectedDataAsync(data[0].title);
      });
      
      


      // fetch(url)
      //   .then(response => response.json())
      //   .then(data => {
      //     console.log(typeof data);
      //     const posts = Object.entries(data);
      //     console.log(posts.title);
      //   })      
    }

    if (selection === "p2")
    {
      docBody.insertParagraph(p2,"End");
    }

    if (selection === "p3")
    {
      docBody.insertParagraph(p3,"End");
    }

    if (selection === "p4")
    {
      docBody.insertParagraph(p4,"End");
    }

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function readtextFile() {
  const response = await fetch('https://jsonplaceholder.typicode.com/todos');
  const data = await response.json();
  
  // results = [];
  // results = data;
  // // data.forEach(element => {
  // //     console.log(element.title)    
  // // });
  
  return data;
}


// async function getData() {

//   let url = 'https://jsonplaceholder.typicode.com/posts';

//   const response = await fetch(url);
//   const data = await response.json();
//   // let titles = new Array();


//   // data.forEach(elt => {
//   //   //console.log(elt.title);
//   //   titles = elt.title;
//   // })

//   docBody.insertParagraph("Hello from getData","End");
// }




// function applyStyle() {
//   Word.run(function (context) {

//     var firstParagraph = context.document.body.paragraphs.getFirst();
//     firstParagraph.styleBuiltIn = Word.Style.intenseReference;

//       return context.sync();
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// function applyCustomStyle() {
//   Word.run(function (context) {

//     var lastParagraph = context.document.body.paragraphs.getLast();
//     lastParagraph.style = "MyCustomStyle";

//       return context.sync();
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// function changeFont() {
//   Word.run(function (context) {

//     var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
//     secondParagraph.font.set({
//             name: "Courier New",
//             bold: true,
//             size: 18
//         });

//       return context.sync();
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// function insertTextIntoRange() {
//   Word.run(function (context) {

//       var doc = context.document;
//       var originalRange = doc.getSelection();
//       originalRange.insertText(" (C2R)", "End");

//       originalRange.load("text");
//       return context.sync()
//           .then(function() {
//               doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
//           })
//           .then(context.sync);
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// function insertTextBeforeRange() {
//   Word.run(function (context) {

//     var doc = context.document;
//     var originalRange = doc.getSelection();
//     originalRange.insertText("Office 2019, ", "Before");

//     originalRange.load("text");
//     return context.sync()
//       .then(function() {
//         doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
//    })
//    .then(context.sync);
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// function replaceText() {
//   Word.run(function (context) {

//     var doc = context.document;
//     var originalRange = doc.getSelection();
//     originalRange.insertText("many", "Replace");

//       return context.sync();
//   })
//   .catch(function (error) {
//       console.log("Error: " + error);
//       if (error instanceof OfficeExtension.Error) {
//           console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       }
//   });
// }

// async function webRequest() {
//   let url = 'https://jsonplaceholder.typicode.com/posts';
//   try {
//       let res = await fetch(url);
//       return await res.json();
//   } catch (error) {
//       console.log(error);
//   }
// }

