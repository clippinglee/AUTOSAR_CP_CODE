

#ifndef _VN7X_H_
#define _VN7X_H_

#include "Vn7x_Types.h"
#include "Vn7x_HwCfg.h"


extern void Vn7x_Init(void);
extern void Vn7x_DeInit(void);
extern void Vn7x_MainFunction(void);
extern void Vn7x_WriteDoChn(uint8 u8Chn, uint16 u16Val);


#endif