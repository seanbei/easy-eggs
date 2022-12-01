/**An example for modbus server/slave using libmodbus(https://www.libmodbus.org/)*/



#include "modbus.h"

int main(int argc, char *argv[])
{
    modbus_t *mb;
    modbus_mapping_t *mb_mapping;
    int server_socket = -1;

    // create tcp channel, set NULL to allow all IPs to visit.
    mb = modbus_new_tcp(NULL, 502);
    mb_mapping = modbus_mapping_new(65535,65535,65535,65535);
    if(mb_mapping == NULL)
    {
        log_error("Failed mapping:%s\n", modbus_strerror(errno));
        modbus_free(mb);
        return -1;
    }

    // set register, here set the index to its value
    int index = 0;
    for(index = 0; index < 65535; index++)
    {
        mb_mapping->tab_registers[index] = index;
    }

    // start to listen
    server_socket = modbus_tcp_listen(mb, 1);
    if(server_socket == -1)
    {
        log_error("Unable to listen TCP.\n");
        modbus_free(mb);
        return -1;  
    }

    modbus_tcp_accept(mb, &server_socket);
    
    while (1)
    {
        uint8_t query[MODBUS_TCP_MAX_ADU_LENGTH];
        int rc;
        rc = modbus_receive(mb, query);
        if(rc > 0)
        {
            modbus_reply(mb, query, rc, mb_mapping);
        }
    }
    
    modbus_mapping_free(mb_mapping);

    modbus_close(mb);
    modbus_free(mb);
    return 0;
}