toxl = require('..');
fs = require('fs');

var test = [
    {
        parent1: {
            child1: {
                child11: 10, 
                child12: 11
            }, 
            child2: {
                child21: 20, 
                child22: 21
            }
        }, 
        parent2 : {
            child1: {
                child11: 101, 
                child21: 102
            }
        }
    }, 
    {
        parent1: {
            child2: {
                child21: 20, 
                child22: 21
            }, 
            child3: {
                child31: 30, 
                child32: 31
            }
        }, 
        parent2 : {
            child1: {
                child11: 101, 
                child21: 102
            }
        }
    }
]

buffer = toxl(test, "My Example", ".");
fs.writeFile("example.xlsx", buffer,  "binary",function(err) {
    if(err) {
        console.log(err);
    } else {
        console.log("Saved excel file to example.xlsx");
    }
});