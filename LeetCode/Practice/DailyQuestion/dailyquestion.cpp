#include "dailyquestion.h"

vector<int> DailyQuestion::nextGreaterElement(vector<int> &nums1, vector<int> &nums2)
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
        ans.push_back(mp[nums2[i]]);
    }
    return ans;
}

int DailyQuestion::minPatches(vector<int> &nums, int n)
{
    // LeetCode 330 : Patching Array
    // Given a sorted integer array nums and an integer n, add/patch elements to the array such that any
    // number in the range [1, n] inclusive can be formed by the sum of some elements in the array.
    // Return the minimum number of patches required.
    // vector<int> store(1000000, 0);
    unordered_set<int> st;
    for (int i = 0; i < nums.size(); i++)
    {
        int temp = 0;
        for (int j = i; j < nums.size(); j++)
        {
            temp += nums[j];
            st.insert(temp);
        }
    }

    int ans = 0;
    for (int i = 1; i <= n; i++)
    {
        cout << endl
             << endl
             << endl
             << " I am at " << i << endl;
        for (auto it : st)
            cout << it << " ";

        if (st.find(i) == st.end())
        {
            cout << "Adding : " << i << endl;
            ans++;
            for (auto it : st)
                st.insert(it + i);
        }
    }

    return ans;
}

int DailyQuestion::findCenter(vector<vector<int>> &edges)
{
    // Leetcode 1791 Find Center of Star Graph
    int siz = edges.size();
    vector<int> store(siz + 2);
    for (int i = 0; i < siz; i++)
    {
        store[edges[i][0]]++;
        store[edges[i][1]]++;
    }
    for (int i = 0; i < store.size(); i++)
    {
        if (store[i] == siz)
            return i;
    }
    return 0;
}

int DailyQuestion::findTheWinner(int n, int k)
{
    // LeetCode 1823. Find the Winner of the Circular Game
    vector<bool> isMarked(n, true);
    int noMarked = 0;
    while (noMarked <= n)
    {
    }
    return 0;
}

int DailyQuestion::minimumDeletions(string s)
{

    class Solution
    {
    public:
        string kthDistinct(vector<string> &arr, int k)
        {
            unordered_map<string, pair<int, int>> mp;
            // for (auto)

            //     for (auto it : arr)
            //     {
            //         mp[it].first++;
            //         mp[it].second =
            //     }
            vector<string> store;
            // for (auto it : mp)
            // {
            //     if (it.second == 1)
            //         store.push_back(it.first);
            // }

            for (auto it : store)
                cout << it << " ";
            if (k > store.size())
                return "";
            return store[k - 1];
        }
    };

    cout << "Testing for the string " << s << endl;

    int ans = 0, count = 0;
    for (int i = 0; i < s.size(); i++)
    {
        if (s[i] == 'b')
            count++;
        else if (count)
        {
            ans++;
            count--;
        }
    }
    cout << ans << endl;
    return ans;
}

int DailyQuestion::maximumBeauty(vector<int> &nums, int k)
{
    int siz = nums.size(), left = 0, right = siz - 1;

    while (left <= right)
    {
        cout << "Printing at indices " << left << " and " << right << ": ";
        cout << nums[left] << ", " << nums[right];
        left++;
        right--;
        cout << endl;
    }

    return 0;
}

int DailyQuestion::countGoodStrings(int low, int high, int zero, int one)
{
    const int MOD = 1e9 + 7;

    return 0;
}

bool isVowel(string s)
{
    int siz = s.size();
    bool flag1 = s[0] == 'a' || s[0] == 'e' || s[0] == 'i' || s[0] == 'o' || s[0] == 'u';
    bool flag2 = s[siz - 1] == 'a' || s[siz - 1] == 'e' || s[siz - 1] == 'i' || s[siz - 1] == 'o' || s[siz - 1] == 'u';
    if (flag1 && flag2)
        return true;
    return false;
}

vector<int> DailyQuestion::vowelStrings(vector<string> &words, vector<vector<int>> &queries)
{
    // LeetCode 2559. Count Vowel Strings in Ranges
    vector<int> ans;
    vector<bool> store(words.size(), false);
    for (int i = 0; i < words.size(); i++)
    {
        if (isVowel(words[i]))
        {
            store[i] = true;
        }
    }
    vector<int> prePorcess(words.size());
    if (store[0])
        prePorcess[0] = 1;
    for (int i = 1; i < words.size(); i++)
    {
        prePorcess[i] = store[i] + prePorcess[i - 1];
    }
    for (int i = 0; i < queries.size(); i++)
    {
        int left = queries[i][0];
        int right = queries[i][1];
        if (left > 0 && right > 0)
        {
            ans.push_back(prePorcess[right] - prePorcess[left - 1]);
        }
        else
        {
            ans.push_back(prePorcess[right]);
        }
    }

    return ans;
}

vector<int> DailyQuestion::findMissingAndRepeatedValues(vector<vector<int>> &grid)
{   
    int siz = grid.size(); 
    int nums = siz*siz; 
    long long sum = 0, sqrSum = 0;
    for(auto it: grid){ 
        for(auto num: it){ 
            sum += num; 
            sqrSum += (long long)(num*num); 
        }
    }

    int sumDiff = sum - (nums*(nums +1)/2); 
    int sqrDiff = sqrSum - nums*(nums + 1)*(2*nums + 1)/6; 

    int repeat = (sqrDiff / sumDiff + sumDiff) / 2;
    int missing = (sqrDiff / sumDiff - sumDiff) / 2;

    return {repeat, missing};
}

int DailyQuestion::longestNiceSubarray(vector<int> &nums)
{   
    //LeetCode 2401. Longest Nice Subarray
    int siz = nums.size(); 
    int ans = 0; 
    for(int i =0; i<siz; i++){ 
        int temp = nums[i]; 
        int j =i+1; 
        while(j< siz && !(nums[j] & nums[i])){ 
            j++; 
        }
        ans = max(ans, j -i); 
    }
    return ans; 
}

long long DailyQuestion::putMarbles(vector<int> &weights, int k)
{   
    int siz = weights.size(); 
    vector<int> store; 
    for(int i = 0; i<siz- 1; i++){ 
        store.push_back(weights[i] + weights[i+1]);
    }
    
    return 0;
}

void DailyQuestion::Test(vector<int> nums)
{   
    cout << "Please provide the inputs"<< endl;
    map<char, pair<int,int>> DIR = {
        {'U',{0,1}}, {'D',{0,-1}},
        {'L',{-1,0}}, {'R',{1,0}}
    };


    vector<pair<char,int>> instr;

    int T; 
    cin >> T; 

    char d;
    int len;
    string color;

    while(T--){
        cin >> d >> len >> color;
        instr.emplace_back(d, len);
    }

    // Trace path, record all dug cells
    set<pair<int,int>> dug;
    int x = 0, y = 0;
    dug.insert({x,y});
    int minx=0, maxx=0, miny=0, maxy=0;

    for (const auto &p : instr) {
        // Explicitly unpack instead of structured binding
        pair<int,int> delta = DIR.at(p.first);
        int dx = delta.first;
        int dy = delta.second;
        for (int i = 0; i < p.second; i++) {
            x += dx;
            y += dy;
            dug.insert({x,y});
            minx = min(minx, x);
            maxx = max(maxx, x);
            miny = min(miny, y);
            maxy = max(maxy, y);
        }
    }

    // Build grid with padding of 1 cell all around
    int W = maxx - minx + 3, H = maxy - miny + 3;
    int ox = -minx + 1, oy = -miny + 1;
    vector<vector<char>> grid(H, vector<char>(W, '.'));
    for (const auto &c : dug) {
        int gx = c.first + ox;
        int gy = c.second + oy;
        grid[gy][gx] = '#';
    }

    // Flood-fill from (0,0) in padded grid to mark exterior
    vector<vector<bool>> vis(H, vector<bool>(W,false));
    queue<pair<int,int>> q;
    q.push({0,0});
    vis[0][0] = true;
    while (!q.empty()) {
        auto cur = q.front(); q.pop();
        int cx = cur.first, cy = cur.second;
        for (const auto &mv : DIR) {
            int nx = cx + mv.second.first;
            int ny = cy + mv.second.second;
            if (nx>=0 && nx<W && ny>=0 && ny<H
                && !vis[ny][nx] && grid[ny][nx]=='.') {
                vis[ny][nx] = true;
                q.push({nx,ny});
            }
        }
    }

    // Count border (#) plus interior cells ('.' not reachable)
    long long total = 0;
    for (int yy = 0; yy < H; yy++) {
        for (int xx = 0; xx < W; xx++) {
            if (grid[yy][xx]=='#' || (grid[yy][xx]=='.' && !vis[yy][xx])) {
                total++;
            }
        }
    }

    cout << total << "\n";
}
