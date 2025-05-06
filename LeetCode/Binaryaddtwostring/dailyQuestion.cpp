#include <iostream>
#include <string>
#include <vector>
#include <unordered_map>
using namespace std;

class Solution
{
public:
    // 3012. Minimize Length of Array Using Operations
    // Your task is to minimize the length of nums by performing the following operations any number of times (including zero):
    // Select two distinct indices i and j from nums, such that nums[i] > 0 and nums[j] > 0.
    // Insert the result of nums[i] % nums[j] at the end of nums.
    // Delete the elements at indices i and j from nums.
    // Return an integer denoting the minimum length of nums after performing the operation any number of times.
    int minimumArrayLength(vector<int> &nums)
    {
        // 
    }

    // Find if Array Can Be Sorted
    int bitsnumber(int n)
    {
        int count = 0;
        while (n)
        {
            if (n & 1)
                count++;
            n >>= 1;
        }
        return count;
    }
    string bubbleSort(vector<int> nums)
    {
        unordered_map<int, int> mp = getBits(nums);
        int siz = nums.size();
        bool swapped;
        for (int i = 0; i < siz - 1; i++)
        {
            swapped = false;
            for (int j = 0; j < siz - i - 1; j++)
            {
                if (nums[j] > nums[j + 1])
                {
                    if (mp.find(nums[j])->second != mp.find(nums[j + 1])->second)
                    {
                        return "Cannt sort the array";
                    }
                    swap(nums[j], nums[j + 1]);
                    swapped = true;
                }
            }
            if (!swapped)
                break;
        }
        return "Congratulations! Your array can be sorted";
    }
    unordered_map<int, int> getBits(vector<int> &nums)
    {
        unordered_map<int, int> mp;
        for (auto it : nums)
        {
            mp[it] = bitsnumber(it);
        }
        // Built in function in GCC compiler to find the number of set bits in a number;
        // for (auto it : nums)
        // {
        //     cout << __builtin_popcount(it) << " ";
        // }
        return mp;
    }
};

int main()
{
    cout << "XF05G^7" << endl;
    vector<int> nums = {3, 16, 8, 4, 2};
    Solution obj;

    return 0;
}