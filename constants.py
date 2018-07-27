LIMIT = 50
TIME_ZONE = "America/New_York"
HTTPCHECKS = ["check-api-umbrella-status-prod","check-chef-cluster","check_http_cfssl","check_http_UCP-replica1-check",
"check_http_UCP-replica2-check","check_http_UCP-controller1-check","check_http_opendj","check_http_openam","check_http_jenkins",
"check_http_marketplace","check-es-cluster-prod","check_http_kibana","check-es-query-count","check-sensu-rabbitmq-cluster",
"check_http_sysdig","check_http_vault","check_http_beta.sam.gov","check_http_alpha.sam.gov","check_http_UCP-replica1-prod-check",
"check_http_UCP-replica2-prod-check","check_http_UCP-ctlr-prod-check","check-openidm-health",
"check-es-cluster-nonprod","check_docker dtr-","check-api-umbrella-status-non-prod"]


#***********EXCEL CONSTANTS***************
inital_date_cell = 'B11'
http_check_row = 8
init_http_check_col = 4
last_index_col = 184
dates_col = 2
init_dates_row = 11
last_index_row = 443
time_increment_row = 2
init_outage_type_col = 8
init_outage_type_row = 13
