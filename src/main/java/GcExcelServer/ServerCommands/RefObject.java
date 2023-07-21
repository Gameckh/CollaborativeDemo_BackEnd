/*
 * Copyright (c) 2019 by GrapeCity, Inc.
 * 
 *  
 * All rights reserved. No part of this source code may be copied, modified 
 * or reproduced in any form without retaining the above copyright notice. 
 * This source code, or source code derived from it, may not be redistributed 
 * without express written permission of GrapeCity, Inc.
 */
package GcExcelServer.ServerCommands;

//----------------------------------------------------------------------------------------
//	Copyright © 2007 - 2017 Tangible Software Solutions Inc.
//	This class can be used by anyone provided that the copyright notice remains intact.
//
//	This class is used to simulate the ability to pass arguments by reference in Java.
//----------------------------------------------------------------------------------------
/**
 * This class is used to simulate interior_ptr&lt;T&gt; in Java.
 */
public final class RefObject<T> {
	public T argValue;

	/**
	 * Wraps a ref (ByRef in Visual Basic) stack object.
	 */
	public RefObject(T refArg) {
		argValue = refArg;
	}

	/**
	 * Wraps an out (&lt;Out&gt; ByRef in Visual Basic) parameter.
	 */
	public RefObject() {

	}
}
