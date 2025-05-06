#include "leetcodecontest.h"

void LeetCodeContest::toTest()
{
    cout << "Hey I am from the leetcode contest" << endl;
}

long long LeetCodeContest::minimumOperations(vector<int> &nums, vector<int> &target)
{
    long long ans = 0;
    int siz = nums.size();
    vector<int> diff(siz + 1);
    for (int i = 0; i < siz; i++)
    {
        diff[i + 1] = target[i] - nums[i];
    }

    for (int i = 0; i < siz; i++)
    {
        if (diff[i + 1] * diff[i] >= 0)
        {
            ans++;
        }
    }

    return ans;
}

int LeetCodeContest::numOfUnplacedFruits(vector<int> &fruits, vector<int> &baskets)
{       
    return 0;
}
