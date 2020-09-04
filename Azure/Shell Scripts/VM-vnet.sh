RGROUP=$(az group create --name vmbackups --location westus2 --output tsv --query name)

az network vnet create \
    --resource-group $RGROUP \
    --name NorthwindInternal \
    --address-prefix 10.0.0.0/16 \
    --subnet-name NorthwindInternal1 \
    --subnet-prefix 10.0.0.0/24

az vm create \
    --resource-group $RGROUP \
    --name NW-APP01 \
    --size Standard_DS1_v2 \
    --vnet-name NorthwindInternal \
    --subnet NorthwindInternal1 \
    --image Win2016Datacenter \
    --admin-username admin123 \
    --no-wait \
    --admin-password <password>

az vm create \
    --resource-group $RGROUP \
    --name NW-RHEL01 \
    --size Standard_DS1_v2 \
    --image RedHat:RHEL:7-RAW:latest \
    --authentication-type ssh \
    --generate-ssh-keys \
    --vnet-name NorthwindInternal \
    --subnet NorthwindInternal1

