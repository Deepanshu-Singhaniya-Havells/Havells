#ifndef LEETCODECONTEST_H
#define LEETCODECONTEST_H

#include <iostream>
#include <vector>
#include <stack>
#include <map>
#include <string>
#include <queue> 
using namespace std;


class TreeNode{ 

    public: 
    int val; 

    TreeNode* left; 
    TreeNode* right; 
    TreeNode(){ 
        val = 0; 
        left =  NULL; 
        right = NULL; 
    }
    TreeNode(int data){ 
        val = data;
        left = NULL; 
        right = NULL; 
    }
};


class LeetCodeContest {
public:
    void toTest(); 

    // Weekly Contest 407 
    long long minimumOperations(vector<int>& nums, vector<int>& target);
    // Weekly Contest 440
    int numOfUnplacedFruits(vector<int>& fruits, vector<int>& baskets);
    // Weekly Context 440 
    

};

#endif //LEETCODECONTEST_H