# stock-analysis

## Overview of Project

### Purpose
The pupose of this challenge was to refactor the code to loop through the stock market dataset one time. Thus making the code more efficient by saving time and using less memory. We want our codes to be as logically sound as possible so we need to weed out unnecessary steps.

## Analysis and Challenges
Biggest challenge I faced: When running my code my For loop kept repeating at every if statement making my ticker value overflow. The solution was to look at my excel spreadsheet, somehow I accidentally changed values in my worksheet that caused my for loop to run incorrectly.

### Results
After refactoring our code it ran much faster than the original. I believe the original code was running around 24 seconds. Whereas, the refactored code ran in microseconds (as pictured below)

![alt text](https://github.com/amarks5/stock-analysis/tree/main/Resources)

## Summary
- What are the advantages or disadvantages of refactoring code?

** Advantages: Refactoring allows you to find bugs in your code. Continously refactoring you can make your code much more clean and organized thus resulting in a faster run time (more efficient code). 
** Disadvantages: Refactoring can be costly in two ways. It can cost you time and money. Not all codes necessarily need to be factored.

- How do these pros and cons apply to refactoring the original VBA script?
Pros: The same as other codes being refactored. In this assignment the refactoring of our code made it run significantly faster and more efficient.  
Cons: I felt the it was more difficult to identify where my errors were in my code.
