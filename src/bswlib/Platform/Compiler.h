#if (!defined COMPILER_H)
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER047_Conf, 1 */
#define COMPILER_H
/**
 * \file
 *
 * \brief AUTOSAR Compiler
 *
 * This file contains the implementation of the AUTOSAR
 * module Compiler.
 *
 * \author Elektrobit Automotive GmbH, 91058 Erlangen, Germany
 *
 * Copyright 2005 - 2018 Elektrobit Automotive GmbH
 * All rights exclusively reserved for Elektrobit Automotive GmbH,
 * unless expressly agreed to otherwise.
 */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER003, 1,
            CompilerAbstraction.ASR403.COMPILER004, 1,
            CompilerAbstraction.ASR403.COMPILER053, 1  */

/*
 * MISRA-C:2012 Deviation List
 *
 * MISRAC2012-1) Deviated Rule: 20.7 (required)
 * Expressions resulting from the expansion of macro parameters shall be
 * enclosed in parentheses.
 *
 * Reason:
 * The AUTOSAR compiler abstraction requires these definitions in this way
 * and the arguments cannot be enclosed in parentheses due to C syntax.
 *
 * MISRAC2012-2) Deviated Rule: 21.1 (required)
 * #define and #undef shall not be used on a reserved identifier or reserved
 * macro name.
 *
 * Reason:
 * Generally, all identifiers that begin with the underscore character are reserved.
 * The macro _INLINE_ is only used for Autosar 2.1 backward compatibility.
 */

/*==================[inclusions]=============================================*/

/* COMPILER052 */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER052_EB, 1 */
#include <Compiler_Cfg.h>    /* module specific memory and pointer classes */
#include <Compiler_Common.h> /* common parts of Compiler.h files */

/*==================[macros]=================================================*/

/*------------------[compiler symbol]----------------------------------------*/

/* COMPILER010 */
/* Also generate processor specific defines as Autosar requirements
   COMPILER010 or COMPILER012 are not clearly stated

   defines can be found in compiler specific folder in Compiler_Specific.h */
#include <Compiler_Specific.h> /* processor specific defines */

#if ( defined COMPILERCFG_EXTENSION_MCAL_FILE )
#include <Compiler_CfgExtMCAL.h> /* MCAL specific defines */
#endif /* if ( defined COMPILERCFG_EXTENSION_MCAL_FILE ) */


/*------------------[storage classes]----------------------------------------*/

#if (defined INLINE)
#error INLINE already defined
#endif

/* COMPILER057 */
/** \brief definition of an inline keyword
 **
 ** To be used for inlining functions
 ** values: 'inline' or empty
 ** to strictly stick to the strict ANSI C90 standard inline is not an option */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER057, 1 */
#define INLINE COMPILER_SPECIFIC_INLINE

#if (defined _INLINE_)
#error _INLINE_ already defined
#endif

/** \brief map _INLINE_ to INLINE for Autosar 2.1 backward compatibility */
/* Deviation MISRAC2012-2 */
#define _INLINE_ INLINE

#if (defined LOCAL_INLINE)
#error LOCAL_INLINE is already defined
#endif

/* COMPILER060 */
/** \brief definition of a keyword for 'static inline' functions
 **
 ** To be used for 'static inline' functions. */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER060, 1 */
#define LOCAL_INLINE static INLINE

/*------------------[macros for functions]-----------------------------------*/

#if (defined FUNC)
#error FUNC already defined
#endif

/* COMPILER001 */
/** \brief abstraction for function declaration and definition
 **
 ** This macro abstracts the declaration and definition of functions
 ** and ensures the correct syntax of function declaration as
 ** required by the specific compiler.
 **
 ** \param[in] rettype return type of the function
 ** \param[in] memclass classification of the function itself */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER001, 1 */
/* Deviation MISRAC2012-1 */
#define FUNC(rettype, memclass) rettype

#if (defined FUNC_P2CONST)
#error FUNC_P2CONST already defined
#endif

/* COMPILER061 */
/** \brief abstraction for function declaration and definition
 **
 ** This macro abstracts the declaration and definition of functions
 ** returning a pointer to a constant and ensures the correct syntax
 ** of function declarations as required by a specific compiler.
 **
 ** \param[in] rettype return type of the function
 ** \param[in] ptrclass defines the classification of the pointers distance
 ** \param[in] memclass classification of the function itself */
/* COMPILER062 */
/** \brief abstraction for function declaration and definition
 **
 ** In the parameter list of the macro FUNC_P2CONST, no further
 ** Compiler Abstraction macros shall be nested. */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER061, 1 ,
            CompilerAbstraction.ASR403.COMPILER062, 1 */
/* Deviation MISRAC2012-1 */
#define FUNC_P2CONST(rettype, ptrclass, memclass) rettype const *

#if (defined FUNC_P2VAR)
#error FUNC_P2VAR already defined
#endif

/* COMPILER063 */
/** \brief abstraction for function declaration and definition
 **
 ** This macro abstracts the declaration and definition of functions
 ** returning a pointer to a variable and ensures the correct syntax
 ** of function declarations as required by a specific compiler.
 **
 ** \param[in] rettype return type of the function
 ** \param[in] ptrclass defines the classification of the pointers distance
 ** \param[in] memclass classification of the function itself */
/* COMPILER064 */
/** \brief abstraction for function declaration and definition
 **
 ** In the parameter list of the macro FUNC_P2VAR, no further
 ** Compiler Abstraction macros shall be nested. */

/* !LINKSTO CompilerAbstraction.ASR403.COMPILER063, 1,
            CompilerAbstraction.ASR403.COMPILER064, 1 */
/* Deviation MISRAC2012-1 */
#define FUNC_P2VAR(rettype, ptrclass, memclass) rettype *

/*------------------[macros for pointers]------------------------------------*/

#if (defined P2VAR)
#error P2VAR already defined
#endif

/* COMPILER006 */
/** \brief abstraction for pointers in RAM pointing to RAM
 **
 ** This macro abstracts the declaration and definition of pointers
 ** in RAM pointing to variables in RAM.
 **
 ** The pointer itself is modifiable.
 ** The pointer's target is modifiable.
 **
 ** \param[in] ptrtype type of the referenced variable
 ** \param[in] memclass classification of the pointer's variable itself
 ** \param[in] defines the classification of the pointer's distance */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER006, 1 */
/* Deviation MISRAC2012-1 */
#define P2VAR(ptrtype, memclass, ptrclass) ptrtype *

#if (defined P2CONST)
#error P2CONST already defined
#endif

/* COMPILER013 */
/** \brief abstraction for pointers in RAM pointing to ROM
 **
 ** This macro abstracts the declaration and definition of pointers
 ** in RAM pointing to constants in ROM.
 **
 ** The pointer itself is modifiable.
 ** The pointer's target is not modifiable (read only).
 **
 ** \param[in] ptrtype type of the referenced constant
 ** \param[in] memclass classification of the pointer's variable itself
 ** \param[in] defines the classification of the pointer's distance */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER013, 1 */
/* Deviation MISRAC2012-1 */
#define P2CONST(ptrtype, memclass, ptrclass) ptrtype const *

#if (defined CONSTP2VAR)
#error CONSTP2VAR already defined
#endif

/* COMPILER031 */
/** \brief abstraction for pointers in ROM pointing to RAM
 **
 ** This macro abstracts the declaration and definition of pointers
 ** in ROM pointing to variables in RAM.
 **
 ** The pointer is not modifiable. (read only).
 ** The pointer's target is modifiable.
 **
 ** \param[in] ptrtype type of the referenced variable
 ** \param[in] memclass classification of the pointer's variable itself
 ** \param[in] defines the classification of the pointer's distance */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER031, 1 */
/* Deviation MISRAC2012-1 */
#define CONSTP2VAR(ptrtype, memclass, ptrclass) ptrtype * const

#if (defined CONSTP2CONST)
#error CONSTP2CONST already defined
#endif

/* COMPILER032 */
/** \brief abstraction for pointers in ROM pointing to ROM
 **
 ** This macro abstracts the declaration and definition of pointers
 ** in ROM pointing to constants in ROM.
 **
 ** The pointer itself is not modifiable (read only).
 ** The pointer's target is not modifiable (read only).
 **
 ** \param[in] ptrtype type of the referenced constant
 ** \param[in] memclass classification of the pointer's variable itself
 ** \param[in] defines the classification of the pointer's distance */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER032, 1 */
/* Deviation MISRAC2012-1 */
#define CONSTP2CONST(ptrtype, memclass, ptrclass) ptrtype const * const

/* COMPILER039 */
#if (defined P2FUNC)
#error P2FUNC already defined
#endif

/** \brief abstraction for declaration and definition of function pointers
 **
 ** This macro abstracts the declaration and definition of pointers
 ** to functions.
 **
 ** \param[in] rettype return type of the function
 ** \param[in] ptrclass defines the classification of the pointer's distance
 ** \param[in] function name respectively name of the defined type
 ** */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER039, 1 */
/* Deviation MISRAC2012-1 */
#define P2FUNC(rettype, ptrclass, fctname) rettype (* fctname)

/* SWS_COMPILER_00065 */
#if (defined CONSTP2FUNC)
#error CONSTP2FUNC already defined
#endif

/** \brief abstraction for declaration and definition of constant function pointers
 **
 ** This macro abstracts the declaration and definition of constant pointers
 ** to functions.
 **
 ** \param[in] rettype return type of the function
 ** \param[in] ptrclass defines the classification of the pointer's distance
 ** \param[in] function name respectively name of the defined type
 ** */
/* Deviation MISRAC2012-1 */
#define CONSTP2FUNC(rettype, ptrclass, fctname) rettype (* const fctname)

/*------------------[keywords for constants]---------------------------------*/

#if (defined CONST)
#error CONST already defined
#endif

/* COMPILER023 */
/** \brief abstraction for declaration and definition of constants
 **
 ** This macro abstracts the declaration and definition of constants.
 **
 ** \param[in] consttype type of the constant
 ** \param[in] memclass classification of the constant itself */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER023, 1 */
/* Deviation MISRAC2012-1 */
#define CONST(consttype, memclass) consttype const

/*------------------[keywords for variables]---------------------------------*/

/* COMPILER026 */
#if (defined VAR)
#error VAR already defined
#endif

/** \brief abstraction for the declaration and definition of variables
 **
 ** This macro abstracts the declaration and definition of variables.
 **
 ** \param[in] vartype type of the variable
 ** \param[in] memclass classification of the variable itself */
/* !LINKSTO CompilerAbstraction.ASR403.COMPILER026, 1 */
/* Deviation MISRAC2012-1 */
#define VAR(vartype, memclass) vartype

/*==================[type definitions]=======================================*/

/*==================[external function declarations]=========================*/

/*==================[internal function declarations]=========================*/

/*==================[external constants]=====================================*/

/*==================[internal constants]=====================================*/

/*==================[external data]==========================================*/

/*==================[internal data]==========================================*/

/*==================[external function definitions]==========================*/

/*==================[internal function definitions]==========================*/

/*==================[end of file]============================================*/
#endif /* if !defined( COMPILER_H ) */
