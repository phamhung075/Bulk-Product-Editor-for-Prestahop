// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
//removeEmptyElements////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Removes null and undefined elements from an array.
 * 
 *  @param {Array} list The array to remove empty elements from
 *  @return {Array} The input array with null and undefined elements removed
*/
const removeEmptyElements = (list) => list.filter((element) => element != null);


//getElementById//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Finds an element by ID within a parent element.
 * 
 *  @param {Element} element The parent element to search within
 *  @param {string} idToFind The ID to search for
 *  @return {Element|null} The element with the specified ID, or null if no such element was found
*/
const getElementById = (element, idToFind) =>
  Array.from(element.getDescendants())
    .map(descendant => descendant.asElement())
    .find(elt => elt?.getAttribute('id')?.getValue() === idToFind);



//getElementsByClassName//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Finds elements by class name within a parent element.
 *  
 *  @param {Element} element The parent element to search within
 *  @param {string} classToFind The class name to search for
 *  @return {Array} An array of elements with the specified class name
*/
const getElementsByClassName = (element, classToFind) =>
  Array.from(element.getDescendants()).concat(element)
    .map(descendant => descendant.asElement())
    .filter(elt => {
      const classes = elt?.getAttribute('class')?.getValue();
      if (classes === classToFind) {
        return true;
      } else if (classes) {
        const classList = classes.split(' ');
        return classList.includes(classToFind);
      } else {
        return false;
      }
    });


//getElementsByAttributeName//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Finds elements by attribute name within a parent element.
 * 
 *  @param {Element} element The parent element to search within
 *  @param {string} attributeName The attribute name to search for
 *  @return {Array} An array of elements with the specified attribute name
*/
const getElementsByAttributeName = (element, AttributeName) => Array.from(element.getDescendants()) .map(descendant => descendant.asElement()) .filter(elt => elt?.getName() === AttributeName);




