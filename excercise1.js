// exercise 1, code a "private count"

const count = (function () {
  let counter = 0;
  return function () {
    counter += 1; return counter
  }
})();
  
console.log(count()); // outputs 0
console.log(count()); // outputs 1
console.log(count()); // outputs 2

// The variable "count" is assigned the return value of a self-invoking function.
// The self-invoking function only runs once. It sets the counter to zero (0), and returns a function expression.