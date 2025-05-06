#include <iostream>
#include "Sorting/sorting.h"
#include "Helper/helper.h"
#include "DailyQuestion/dailyquestion.h"
#include "LeetCodeContest/leetcodecontest.h"
#include "CodeForces/codeforces.h"

using namespace std;

int main()
{
    cout << "Start of the Main function" << endl << endl;
    vector<int> nums = {1, 3, 5, 1, 7, 4,2, 6,7, 5};
    vector<vector<int>> edges = {{1,2},{5,1},{1,3}, {1,4}}; 
    
    string str = "bbaaaaabb";

    Sorting sortObj;
    Helper helperObj;
    DailyQuestion dQ;
    LeetCodeContest contest;

    dQ.Test(nums);
    
    //cout << "Center is " << dQ.findCenter(edges) << endl; 
    //cout << endl<< "Min Patches : " << dQO.minPatches(nums, n) << endl;
    // contest.toTest();
    // helperObj.printVector(sortObj.selectionSort(nums));
    // helperObj.printVector(nums);
}
