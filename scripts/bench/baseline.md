# Scaling baseline

Measured on Apple M4 with 16 GiB RAM, macOS 26.5.1, and Node.js 22.18.0. Times are milliseconds; ratios compare each size with the preceding row.

| Case | Paragraphs | Import | State | Edit | Export |
| --- | ---: | ---: | ---: | ---: | ---: |
| plain | 500 | 158.7 | 58.6 | 2.3 | 145.8 |
| plain | 1,000 | 216.1 (1.36x) | 103.8 (1.77x) | 6.1 (2.69x) | 539.0 (3.70x) |
| plain | 2,000 | 367.1 (1.70x) | 220.5 (2.12x) | 21.0 (3.43x) | 2,512.7 (4.66x) |
| plain | 4,000 | 793.9 (2.16x) | 535.6 (2.43x) | 78.0 (3.72x) | 18,306.7 (7.29x) |
| rich | 500 | 114.2 | 53.4 | 2.0 | 331.5 |
| rich | 1,000 | 224.0 (1.96x) | 107.2 (2.01x) | 6.8 (3.46x) | 1,205.0 (3.63x) |
| rich | 2,000 | 462.2 (2.06x) | 238.3 (2.22x) | 25.5 (3.72x) | 6,595.9 (5.47x) |
| rich | 4,000 | 1,070.0 (2.31x) | 572.4 (2.40x) | 100.1 (3.93x) | 82,056.2 (12.44x) |
