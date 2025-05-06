#ifndef DAILYQUESTION_H
#define DAILYQUESTION_H

#include <iostream>
#include <vector>
#include <stack>
#include <map>
#include <set>
#include <queue>
#include <unordered_set>
#include <unordered_map>
using namespace std;

// class TreeNode{
//     int val;
//     TreeNode *left;
//     TreeNode *right;
//     TreeNode() : val(0), left(nullptr), right(nullptr){}
//     TreeNode(int x) : val(x), left(nullptr), right(nullptr) {}
//     TreeNode(int x, TreeNode *left, TreeNode *right) : val(x), left(left), right(right){}
// };

class DailyQuestion
{
public:
    vector<int> nextGreaterElement(vector<int> &nums1, vector<int> &nums2);
    int minPatches(vector<int> &nums, int n);
    int findCenter(vector<vector<int>> &edges);
    int findTheWinner(int n, int k);
    int minimumDeletions(string s);
    int maximumBeauty(vector<int> &nums, int k);
    int countGoodStrings(int low, int high, int zero, int one);
    vector<int> vowelStrings(vector<string> &words, vector<vector<int>> &queries);
    vector<int> findMissingAndRepeatedValues(vector<vector<int>> &grid);
    int longestNiceSubarray(vector<int> &nums);
    long long putMarbles(vector<int> &weights, int k);
    void Test(vector<int> nums);
};

#endif // DAILYQUESTION_H
