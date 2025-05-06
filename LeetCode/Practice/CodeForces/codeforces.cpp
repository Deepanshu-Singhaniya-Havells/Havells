#include <iostream>
#include <codeforces.h>

void solve() { 
    int n; 
    cin >> n; 

    vector<int> store(n, 0); 

    while(n--){ 
        int temp; 
        cin >> temp; 
        store.push_back(temp); 
    }
}

int main()
{
    int T; 
    cin >> T; 
    while(T--){ 
        solve(); 
    }

    return 0; 
}


