metrics = {
    "Linux by Zabbix agent": {
        "Saturation": [
            {"name": "CPU Utilization(%)", "key": "system.cpu.util"},
            {"name": "Memory Utilization(%)", "key": "vm.memory.utilization"},
            {"name": "Number of CPUs", "key": "system.cpu.num"},
            {"name": "Available memory(%)", "key": "vm.memory.size[pavailable]"},
            {"name": "Free swap space", "key": "system.swap.size[,free]"},
            {"name": "Free swap space in %", "key": "system.swap.size[,pfree]"},
            {"name": "Space: Available", "key": 'vfs.fs.dependent.size[/*,free]'},
            {"name": "Space: Used", "key": 'vfs.fs.dependent.size[/*,used]'},
            {"name": "Space: Total", "key": 'vfs.fs.dependent.size[/*,total]'},
            {"name": "Space: Used, in %", "key": 'vfs.fs.dependent.size[/*,pused]'},
            {"name": "Total swap space", "key": "system.swap.size[,total]"},
            {"name": "Number of processes", "key": "proc.num"},
            {"name": "Number of running processes", "key": "proc.num[,,run]"},
            {"name": "Maximum number of processes", "key": "kernel.maxproc"},
            {"name": "Maximum number of open file descriptors", "key": "kernel.maxfiles"},
        ],
        "Traffic": [
            {"name": "Bits received", "key": 'net.if.in[\"e*\"]'},
            {"name": "Bits sent", "key": 'net.if.out[\"e*\"]'},
            {"name": "Speed", "key": 'vfs.file.contents["/sys/class/net/enp1s0/speed"]'},
        ],
        "Latency": [
            {"name": "Load average (1m avg)", "key": 'system.cpu.load[all,avg1]'},
            {"name": "Load average (5m avg)", "key": 'system.cpu.load[all,avg5]'},
            {"name": "Load average (15m avg)", "key": 'system.cpu.load[all,avg15]'},
        ],
        "Errors": [
            {"name": "Inbound packets discarded", "key": 'net.if.in[\"e*\",dropped]'},
            {"name": "Outbound packets discarded", "key": 'net.if.out[\"e*\",dropped]'},
            {"name": "Inbound packets with errors", "key": 'net.if.in[\"e*\",errors]'},
            {"name": "Outbound packets with errors", "key": 'net.if.out[\"e*\",errors]'},
        ],

    }
}