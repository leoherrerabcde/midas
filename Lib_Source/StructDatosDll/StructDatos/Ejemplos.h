/*		if ( EstadoArchivoReproduc )
		{
			GV_Num_Adq_Fifo				++;
			do
			{
				lv_Data_Leida			= _read ( HandleArchivoPulsosReproduc , lv_str_Encabezado , 1 );
				if ( lv_Data_Leida < 1 )
				{
					_close ( HandleArchivoPulsosReproduc );
					EstadoArchivoReproduc = 0;
					return 0 ;
				}
			}
			while ( * lv_str_Encabezado != 'L' );

			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , & ( lv_str_Encabezado [1] ), 8				);
			if ( !lv_Data_Leida )
			{
				if ( pt_Lista_Archivos_Now != NULL )
				{
					_close ( HandleArchivoPulsosReproduc );
					HandleArchivoPulsosReproduc		= _open ( pt_Lista_Archivos_Now->NombreArchivo , _O_BINARY );
					/*if ( GV_Archivo_Actual_Reproduccion )
					{
						free ( GV_Archivo_Actual_Reproduccion );					
					}*/
					//GV_Archivo_Actual_Reproduccion	= ( char * ) malloc ( strlen ( pt_Lista_Archivos_Now->NombreArchivo ) + 1 ) ;
					//memcpy ( GV_Archivo_Actual_Reproduccion , pt_Lista_Archivos_Now->NombreArchivo , strlen ( pt_Lista_Archivos_Now->NombreArchivo ) + 1 );
					GV_Archivo_Actual_Reproduccion	= pt_Lista_Archivos_Now->NombreArchivo ;
					pt_Lista_Archivos_Now			= pt_Lista_Archivos_Now->pt_Next ;
					lv_Data_Leida					= _read ( HandleArchivoPulsosReproduc , lv_str_Encabezado , 9				);
					GV_Num_Adq_Fifo					= 1;
				}
				else
				{
					lv_b_ReproducirPulsos = 0;		// Fin Reproducción.
					return 0;
				}
			}
			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , & dwWordsDisp     , 4				);

			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , lv_str_Encabezado, 7 );
			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , & lv_Tiempo , sizeof ( struct_Time.time ) );

			GV_Tiempo_Ejecucion			= lv_Tiempo;

			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , lv_str_Encabezado, 4 );
			lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , & lv_mili_sec , sizeof ( struct_Time.millitm ) );
			//lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , lv_str_Encabezado, 2 );

			// Calcular Retardo para simular Tiempo Real
			//lv_Time_Now					= GetTickCount ();
			//lv_Delta_mili_seg_TpoReal	= lv_Time_Now - PV_Time_Old;
			PV_Delta_Reprod		= ( lv_Tiempo - PV_Tiempo_Old ) * 1000 + ( lv_mili_sec - PV_mili_sec_Old );
			if ( PV_Delta_Reprod	> 3000 )
			{
				PV_Delta_Reprod		=PV_Delta_Reprod		;
			}
			/*if ( lv_Delta_mili_seg_Repro > lv_Delta_mili_seg_TpoReal )
			{
				if ( lv_Delta_mili_seg_Repro - lv_Delta_mili_seg_TpoReal < 1000 )
				{
					Sleep ( lv_Delta_mili_seg_Repro - lv_Delta_mili_seg_TpoReal );
				}
				PV_Delta_Espera	= lv_Delta_mili_seg_Repro - lv_Delta_mili_seg_TpoReal;
			}
			else
			{
				PV_Delta_Espera	= 0;
			}*/
			PV_Tiempo_Old	= lv_Tiempo ;
			PV_mili_sec_Old	= lv_mili_sec;

			GV_Tick_Count_Repro	= lv_Tiempo * 1000 + lv_mili_sec;

			//PV_Time_Old		= lv_Time_Now;

			if ( dwWordsDisp <= 512 * 32 )
			{
				pulReadBuffer = diriniReadBuffer ;
				if (! pulReadBuffer)
				{
					return 0;
				}
				lv_Data_Leida				= _read ( HandleArchivoPulsosReproduc , pulReadBuffer     , 4 * dwWordsDisp	);
				
				// Corregir Errores del Archivo.
				ptr_Buffer_Fifo			= ( unsigned char * ) pulReadBuffer;
				ptr_Buff_Dest			= ( Buff_Fifo * ) pulReadBuffer;
				ptr_Inicio_Data_Ok		= NULL;
				//dwWordsDisp				= 0;
				lv_Count_Bytes			= 0;
				lv_Size_Data_Ok			= 0;
				lv_bool_Desincronizacion= 0;

				//Corregir_Buffer_Leido ( ptr_Buffer_Fifo , lv_Data_Leida , HandleArchivoPulsosReproduc );

				while ( ( lv_Count_Bytes < lv_Data_Leida ) && ( lv_Data_Leida - lv_Count_Bytes  >= sizeof ( Buff_Fifo ) ) )
				{
					lv_Bytes_Lectura	= ( unsigned long ) Analizar_CR_LF ( ptr_Buffer_Fifo , lv_Data_Leida - lv_Count_Bytes );
					if ( lv_Bytes_Lectura )
					{
						if ( lv_Bytes_Lectura != ( unsigned long ) _read ( HandleArchivoPulsosReproduc , ( unsigned char * ) pulReadBuffer + lv_Data_Leida - lv_Bytes_Lectura , lv_Bytes_Lectura	) )
						{
							lv_Bytes_Lectura = lv_Bytes_Lectura ;
						}
					}
					
					if ( ( ( ( Buff_Fifo * ) ptr_Buffer_Fifo )->cont1 == 0 ) &&
						 ( ( ( Buff_Fifo * ) ptr_Buffer_Fifo )->cont2 == 1 ) &&
						 ( ( ( Buff_Fifo * ) ptr_Buffer_Fifo )->cont3 == 2 ) &&
						 ( ( ( Buff_Fifo * ) ptr_Buffer_Fifo )->cont4 == 3 ) )
					{
						// Data Ok.
						// Verifica si se inicializó Puntero al primer dato ok de la Fifo.
						if ( ptr_Inicio_Data_Ok == NULL )
						{
							ptr_Inicio_Data_Ok = ptr_Buffer_Fifo ;
						}
						// Incrementa puntero y contador de tamaño de data.
						lv_Size_Data_Ok ++;
						/*if ( ( ( Buff_Fifo * ) ptr_Buffer_Fifo )->toa > 85720000 )
						{
							lv_Count_Bytes =lv_Count_Bytes ;
						}*/
						ptr_Buffer_Fifo	+= sizeof ( Buff_Fifo );
						lv_Count_Bytes  += sizeof ( Buff_Fifo );
					}
					else
					{
						// Se desincronizó.
						lv_bool_Desincronizacion	= 1;
						// Verificar si se debe transferir data ok.
						if ( ptr_Inicio_Data_Ok != NULL )
						{
							//lv_Size_Data_Ok		= ptr_Buffer_Fifo - ptr_Inicio_Data_Ok ;
							memcpy ( ptr_Buff_Dest , ptr_Inicio_Data_Ok , ptr_Buffer_Fifo - ptr_Inicio_Data_Ok );//lv_Size_Data_Ok ;
							// Actualizar puntero destino.
							ptr_Buff_Dest		+= lv_Size_Data_Ok;
							// Resetear puntero.
							ptr_Inicio_Data_Ok	= NULL ;
							lv_Size_Data_Ok		= 0;
						}
						// Apuntar al siguiente Byte.
						ptr_Buffer_Fifo ++;
						lv_Count_Bytes ++;
					}
				}
				//
				// Verificar si hubo desincronización.
				if ( lv_bool_Desincronizacion )
				{
					// Verificar si hay que transferir un bloque de data sincronizada.
					if ( ptr_Inicio_Data_Ok != NULL )
					{
						//lv_Size_Data_Ok		= ptr_Buffer_Fifo - ptr_Inicio_Data_Ok ;
						memcpy ( ptr_Buff_Dest , ptr_Inicio_Data_Ok , ptr_Buffer_Fifo - ptr_Inicio_Data_Ok );//lv_Size_Data_Ok ;
						// Actualizar puntero destino.
						ptr_Buff_Dest		+= lv_Size_Data_Ok;
						// Resetear puntero.
						ptr_Inicio_Data_Ok	= NULL ;
						//lv_Size_Data_Ok		= 0;
					}
					// Medir los datos sincronizados.
					dwWordsDisp				= ( ptr_Buff_Dest - ( Buff_Fifo * ) pulReadBuffer ) << 2;
				}
			}
			else
			{
				return 0;
			}
		}
		else
		{
			if ( pt_Lista_Archivos_Now != NULL )
			{
				HandleArchivoPulsosReproduc		= _open ( pt_Lista_Archivos_Now->NombreArchivo , _O_BINARY );
				if ( HandleArchivoPulsosReproduc == -1 )
				{
					lv_b_ReproducirPulsos = 0;		// Fin Reproducción.
					return 0;
				}
				/*if ( GV_Archivo_Actual_Reproduccion )
				{
					free ( GV_Archivo_Actual_Reproduccion );
				}*/
				//GV_Archivo_Actual_Reproduccion	= ( char * ) malloc ( strlen ( pt_Lista_Archivos_Now->NombreArchivo ) + 1 ) ;
				//memcpy ( GV_Archivo_Actual_Reproduccion , pt_Lista_Archivos_Now->NombreArchivo , strlen ( pt_Lista_Archivos_Now->NombreArchivo ) + 1 );
				GV_Archivo_Actual_Reproduccion	= pt_Lista_Archivos_Now->NombreArchivo ;
				EstadoArchivoReproduc			= 1;
				/*if ( ( lv_Tpo_Ejecucion = Get_Hora_From_File ( pt_Lista_Archivos_Now->NombreArchivo ) ) )
				{
					GV_Tiempo_Ejecucion		= lv_Tpo_Ejecucion ;	//			= Get_Hora_From_File ( pt_Lista_Archivos_Now->NombreArchivo );
				}
				else
				{
					lv_Tpo_Ejecucion=lv_Tpo_Ejecucion;
				}*/
				pt_Lista_Archivos_Now	= pt_Lista_Archivos_Now->pt_Next ;
				GV_Num_Adq_Fifo					= 1;
			}
			else
			{
				lv_b_ReproducirPulsos = 0;		// Fin Reproducción.
				return 0;
			}
			return 0;
		}
		//return ( ( int ) dwWordsDisp );		// No hace bada más.
	}*/
