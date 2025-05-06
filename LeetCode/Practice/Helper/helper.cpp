#include "helper.h"
#include <iostream>
void Helper::printVector(vector<int> nums)
{
    cout << "Printing the vector" << endl;
    for (auto it : nums)cout << it << " ";
    cout << endl << "Finished Printing" << endl;
}