import numpy as np
import matplotlib.pyplot as plt  # 画图用
import scipy.io as sci  # .mat文件读取用

filePath = 'C:/Users/Abyss/Desktop/test1.mat'
data = sci.loadmat(filePath)

fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

x = data['xx']
y = data['yy']
z = data['zz']
c = data['cc']
print(np.array(x).shape)
x = x.T  # 不转置也没影响
y = y.T
z = z.T
c = c.T
print(np.array(x).shape)

ax.scatter(x, y, z, c=c, marker='.')
plt.show()
