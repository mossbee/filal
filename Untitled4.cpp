#include <iostream> 
#include <fstream>  
#include <iomanip>  
#include <cmath>    

// Neu so co tri tuyet doi nho hon 1e^-10 thi co the coi la 0
#define ISZERO(x) (fabs((x)) < 1e-10)

using namespace std;

int m, n;
double **a;
const char *checkfile;  

// Tai ma tran tu file
void loadMatrix(const char *pathtofile) {
    ifstream fi(pathtofile);
    checkfile = pathtofile;
    
    fi >> m >> n;
    a = new double *[m + 1]; // Ma tran A dung de luu ma tran bo sung A|B. Cot O cua ma tran A se luu so so 0 o dau moi hang
    for (int i = 0; i <= m; i++) a[i] = new double[n + 2]();
    for (int i = 1; i <= m; i++) for (int j = 1; j <= n + 1; j++) fi >> a[i][j];
    fi.close();
}

// In ma tran ra man hinh
void inmatran(int precision = 4) {
    for (int i = 1; i <= m; i++) {
        cout << (int)a[i][0] << " : "; // a[i][0] dung de luu so so 0 dau tien o hang i
        for (int j = 1; j <= n + 1; j++)
            cout << setw(precision + 5) << left << setprecision(precision) << fixed << a[i][j] << "  ";
        cout << endl;
    }
    cout << "----------------------------------------------------------------" << endl;
}

// Dao 2 hang i,j (i,j khac nhau)
void daohang(int i, int j) {
    if (i != j) swap(a[i], a[j]);
}

// Dem so so 0 dau o moi hang tu hang i tro xuong
void countzeros(int i = 1) {
    for (int k = i; k <= m; k++) {
        int s = 1;
        a[k][0] = 0; 
        while (ISZERO(a[k][s]) && s <= n) a[k][0] = s++; 
    }
}

// \\ Sap xep lai thu tu cac hang tu hang i theo tang dan so so 0 o dau bang sap xep lua chon
void sapxeplaihang(int i = 1) {
    int MIN;
    countzeros(i); 
    for (int z = i; z < m; z++) {
        MIN = z;
        for (int k = z + 1; k <= m; k++) if (a[k][0] < a[MIN][0]) MIN = k;
        daohang(MIN, z);
    }
}

// Bien doi ve ma tran bac thang
void matranbacthang() {
    int i = 1, j = 1;
    while (ISZERO(a[i][j])) {
        i++;
        if (i == m + 1) {
            i = 1; j++;
            if (j == n + 2) {
                countzeros();
                return;
            } 
        }
    }
    cout << "Vi tri dau tien khac 0: " << i << " " << j << endl;
    if (i != 1) daohang(i, 1);

    // Bat dau lap tu cot thu j vua tim duoc o tren den n
    for (; j <= n + 1; j++) {
        for (i = 2; i <= m; i++) {
        
            if (!ISZERO(a[i][j])) {
                int k;
                bool flag = false;
                for (k = i - 1; k >= 1; k--) { 
                    if (!ISZERO(a[k][j]) && a[k][0] == j - 1) {
                        flag = true; 
                        break;
                    }
                }
                if (!flag) continue;

        
                double ratio = a[i][j] / a[k][j];
                for (int z = j; z <= n + 1; z++) a[i][z] = a[i][z] - ratio * a[k][z];
                inmatran(); 
            }
        }
        // Co the khong can thiet sap xep lai, ta dao hang trong luc khu, hien tai chua co cach lam 
        sapxeplaihang();
    }
}
// Xac dinh he co nghiem duy nhat, vo so nghiem hay vo nghiem
int xacdinhloainghiem() {
    int rankA = m, rankAn = m; 
    for (int i = m; i >= 1; i--) {
        if (a[i][0] == n) {    
            rankA--;
            if (ISZERO(a[i][n + 1])) rankAn--; 
        }
        else break;
    }

    if (rankA == rankAn && rankA < n)  return rankA; 
    if (rankA == rankAn && rankA == n) return -1;    
    return -2; 
}

// Kiem tra nghiem trong truong hop he co nghiem duy nhat
bool kiemtra(double *x, int offset = 0, bool param = false) {
    ifstream fi(checkfile);
    fi >> m >> n;

    for (int i = 0; i < m; i++) {
        double tmp = 0.0, aij = 0.0;
        for (int j = 0; j < n; j++) {
            fi >> aij;
            tmp += aij * x[j + 1 + offset];
        }
        fi >> aij;
        if ((param && !ISZERO(tmp)) || (!param && fabs(tmp - aij) > 1e-6)) return false;
    }
    fi.close();
    return true;
}

// Giai nghiem trong truong hop co nghiem duy nhat
void giainghiemduynhat() {
    double *x = new double[n + 1]();
    for (int i = n; i >= 1; i--) {
        x[i] = a[i][n + 1];
        for (int j = n; j > i; j--) x[i] -= x[j] * a[i][j];
        x[i] /= a[i][i];
    }

    cout << "He co nghiem duy nhat" << endl;
    for (int i = 1; i <= n; i++) cout << "x_" << i << " = " << x[i] << endl;
    if (kiemtra(x)) cout << endl << "---> True" << endl;
    else            cout << endl << "---> False" << endl;
};

// Giai nghiem trong truong hop vo so nghiem
void giainghiemtongquat(int rankA) {
    double **x;

    // Tao mot ma tran luu nghiem
    x = new double *[n+1];
    for (int i = 0; i <= n; i++) x[i] = new double[n]();
    for (int i = 0; i < n; i++)  x[i + 1][i] = 1;
    cout << "Rank A: " << rankA << endl;

    // Bieu dien an nay qua an khac (x1 = 0 -2x2 - 3x3 - 3.5x4)
    int i = rankA;
    while (i >= 1) {
        int j = a[i][0], k = 0;
        double aij = a[i][j + 1];
        for (int z = 0; z <= n; z++) x[z][j] = 0;
        x[0][j] = a[i][n + 1] / aij;
        while (++k <= n) if (k != j + 1) x[k][j] = -a[i][k] / aij;
        i--;
    }

    // The an nay vao an kia vi xi bieu dien qua cac xj voi j>i ne se mat dan cac tham so
    for (int i = 0; i < n; i++)
        for (int j = i + 2; j <= n; j++) // Vi du the cac xj khac vao x2 thi j>2
            if (!ISZERO(x[j][i])) {
                double tmp = x[j][i]; // he so cua xj trong bieu dien xi tu buoc tren (xi = ... -2xj + ... th? tmp = -2)
                for (int z = 0; z <= n; z++) x[z][i] = x[z][i] + tmp * x[z][j - 1]; // The xj vao xi
                x[j][i] = x[j][i] - tmp; 
            }

    // In nghiem
    for (int i = 0; i < n; i++) {
        cout << "x_" << i + 1 << " = " << x[0][i];
        for (int j = 1; j <= n; j++) if (!ISZERO(x[j][i]))
                cout << " " << showpos << x[j][i] << noshowpos << "x_" << j;
        cout << endl;
    }

    // Kiem tra lai nghiem
    for (int i = 0; i <= n; i++) {
        if (!kiemtra(x[i], -1, i)) {
        	cout << endl << "---> False" << endl;
            break;
        }
        else if (i==n) cout << endl << "---> True" << endl;
    }
}

// Phuong phap Gauss giai he pt
void Gauss() {
    matranbacthang();
    inmatran();
    int r = xacdinhloainghiem();
    if (r == -2) cout << "He vo nghiem" << endl;
    else if (r == -1)    giainghiemduynhat();
    else                 giainghiemtongquat(r);
}

int main() {
    loadMatrix("input_gauss.txt");
    Gauss();
    return 0;
}
