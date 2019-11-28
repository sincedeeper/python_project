import psutil

print("CPU 数量:{0}".format(psutil.cpu_count()))
print("CPU 数量-不含逻辑core:{0}".format(psutil.cpu_count(logical=False)))
print("物理内存:{0}".format(psutil.virtual_memory()))
print("虚拟内存:{0}".format(psutil.swap_memory()))
print("磁盘分区:{0}".format(psutil.disk_partitions()))
print("网络接口:{0}".format(psutil.net_if_addrs()))
