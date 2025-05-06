// Practice VS.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <vector>
#include <stack>
#include <map>
#include <algorithm>
#include <queue>

using namespace std;

class MyStack {
public:
    queue<int> qu; 
    queue<int> temp;


    MyStack() {
        
    }
    
    void push(int x) {
        for(int i  =0; i<qu.size(); i++){
            temp.push(qu.front());
            qu.pop();
        }
        qu.push(x);

        for(int i  =0; i<temp.size(); i++){
            qu.push(temp.front());
            temp.pop();
        }
    }
    
    int popu() {
        int temp = qu.front(); 
        qu.pop();
        return temp;
    }
    
    int top() {
     return qu.front();   
    }
    
    bool empty() {
        return qu.empty();
    }
};



class Nmeetings
{
    public:


    
    struct meeting
    {
        int start;
        int end;
        int pos;
    };

    static bool comparator(struct meeting meet1, struct meeting meet2)
    {
        if (meet1.end < meet2.end)
            return true;
        else if (meet1.end > meet2.end)
            return false;
        else if (meet1.pos < meet2.pos)
            return true;
        return false;
    }

     int maxMeetings(int start[], int end[], int n)
    {
        struct meeting meet[10001];
        for (int i = 0; i < n; i++)
        {
            meet[i].start = start[i];
            meet[i].end = end[i];
            meet[i].pos = i;
        }

        sort(meet, meet + n, comparator);

        vector<int> ans;

        int limit = meet[0].end;
        ans.push_back(meet[0].pos);
        for (int i = 1; i < n; i++)
        {
            if (meet[i].start > limit)
            {
                limit = meet[i].end;
                ans.push_back(meet[i].pos);
            }
            return ans.size();
        }
    }
};

vector<int> nextGreaterElement(vector<int> &nums1, vector<int> &nums2)
{
    int s1 = nums1.size();
    int s2 = nums2.size();
    stack<int> s;
    map<int, int> mp;

    for (int i = s2 - 1; i > 0; i--)
    {

        if (s.empty())
        {
            mp[nums2[i]] = -1;
        }
        else
        {
            while (s.top() <= nums2[i])
            {
                if (s.empty())
                {
                    mp[nums2[i]] = -1;
                    break;
                }
                s.pop();
            }
            mp[nums2[i]] = s.top();
        }
        s.push(nums2[i]);
    }

    vector<int> ans;
    for (int i = 0; i < nums1.size(); i++)
    {
        ans.push_back(mp[nums1[i]]);
        cout << nums1[i] + " : " + mp[nums1[i]] << endl;
    }

    for (auto it : ans)
    {
        cout << it << endl;
    }
    return ans;
}

int main()
{

    cout << "hello world" << endl;

    MyStack stac; 
    stac.push(1); 
    stac.push(2);
    stac.push(3);
    stac.push(4);
    cout << stac.top() << endl;
    
    //vector<int> nums1 = {4, 1, 2};
    //vector<int> nums2 = {1, 2, 3, 4};

    //// nextGreaterElement(nums1, nums2);

    //int start[6] = {1, 0, 3 , 5, 8, 5};
    //int end[6] = {2, 6, 4, 7, 9, 9};

    //int n = 6;

    // Nmeetings ob; 
    // cout << ob.maxMeetings(start, end, n);



}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started:
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
