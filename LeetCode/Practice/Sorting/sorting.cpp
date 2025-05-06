#include "sorting.h"

vector<int> Sorting::selectionSort(vector<int> nums) {
    int size = nums.size();
    for (int i = 0; i < size; ++i) {
        int min_index = i;
        for (int j = i + 1; j < size; ++j) {
            if (nums[j] < nums[min_index]) {
                min_index = j;
            }
        }
        swap(nums[i], nums[min_index]);
    }
    return nums;
}
