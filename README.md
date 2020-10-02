# Excercises required to apply to work oportunity by "factorconecta.com" 

#### Exercise 1

Code a `count()` function in Javascript.

This function should return an integer in the following fashion:

```js
const count = undefined // replace this with your code;

console.log(count()); // outputs 0
console.log(count()); // outputs 1
console.log(count()); // outputs 2
```

In other words, the function should 'remember' how many times it has been invoked.

Global variables are forbidden. Provide a brief explanation of how your solution works.

#### Exercise 2

Sample API in Javascript.

Create an HTTP service (you are free to use libraries available in NPM) with a single route `POST /palindrome`.

The request to `POST /palindrome` should accept a JSON payload with the following shape:

```js
{
  "phrase": string
}
```

Your service should return a boolean value indicating if the phrase is a palindrome or not ([Wikipedia](https://es.wikipedia.org/wiki/Pal%C3%ADndromo)):

```js
{
  "palindrome": boolean
}
```

**Note**: The code to determine if the phrase is a palindrome should be written by you (not using a library).

#### Exercise 3

Counter in React or Vue

Create a button click counter:

- It should have a 'Count' button.
- It should have a text with the number of times the button has been clicked.
- The text should be inmmediately updated everytime the user clicks on the button.
