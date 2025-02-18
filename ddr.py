import webbrowser

# ddr_start = 
# ddr_end = 
# coll_start = 
# coll_end = 

ddr= 'https://collections.api.mintifi.com/api/ddr/v1/drawdown_requests/download.csv?Authorization=Bearer%20eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiI4NTEiLCJzY3AiOiJhZG1pbl91c2VyIiwiYXVkIjpudWxsLCJpYXQiOjE3MTQ0MDQ1ODgsImV4cCI6MTcxNDQxMDU4OCwianRpIjoiYjUzZTU2NzUtMTg0YS00NTU3LWE3MDctZjgyNGViNzMwMDcxIn0.MQWJjHv54Yp-AsoFN_HYJCP0mPBz5qSZaBhhwJROiKA&q[created_at_gteq]=2024-04-29&q[created_at_end_of_day_lteq]=2024-04-29'  

coll = 'https://collections.api.mintifi.com/api/ddr/v1/payment_distributions.csv?Authorization=Bearer%20eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiI4NTEiLCJzY3AiOiJhZG1pbl91c2VyIiwiYXVkIjpudWxsLCJpYXQiOjE3MTQ0MDQ1ODgsImV4cCI6MTcxNDQxMDU4OCwianRpIjoiYjUzZTU2NzUtMTg0YS00NTU3LWE3MDctZjgyNGViNzMwMDcxIn0.MQWJjHv54Yp-AsoFN_HYJCP0mPBz5qSZaBhhwJROiKA&q[created_at_gteq]=2024-04-29&q[created_at_end_of_day_lteq]=2024-04-29'

webbrowser.open_new_tab(ddr)
webbrowser.open_new_tab(coll)