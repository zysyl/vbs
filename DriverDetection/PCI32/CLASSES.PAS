{ PCI Classes }

type pci_record =
  record
    clase,
    subclass,
    progif      : byte;
    name        : string[40];
  end;

const
   pci_class_names      : array [0..$11] of string[40] = (

   'Unknown',
   'Storage',
   'Network',
   'Display',
   'Multimedia',
   'Memory',
   'Bridge',
   'Simple Communication',
   'System',
   'Input',
   'Docking',
   'Processor',
   'Serial',
   'Wireless',
   'Intelligent I/O',
   'Satellite Communication',
   'En/Decryption',
   'Data Aquisition & Signal Processing'
   );


   high_class_name	: word = $11;	{ update this to same as line above }

   high_class_array	: word = 145;	{ update this to same as line below }

   pci_class_array : array [0..145] of pci_record = (

   (clase:$00;subclass:$00;progif:$00;name:'Undefined'),
   (clase:$00;subclass:$01;progif:$00;name:'VGA'),

   (clase:$01;subclass:$00;progif:$00;name:'SCSI'),
   (clase:$01;subclass:$01;progif:$00;name:'IDE (ATA)'),
   (clase:$01;subclass:$02;progif:$00;name:'Floppy'),
   (clase:$01;subclass:$03;progif:$00;name:'IPI Bus'),
   (clase:$01;subclass:$04;progif:$00;name:'RAID'),
   (clase:$01;subclass:$05;progif:$20;name:'ATA (Single ADMA)'),
   (clase:$01;subclass:$05;progif:$30;name:'ATA (Continuous ADMA)'),
   (clase:$01;subclass:$06;progif:$00;name:'Serial ATA'),
   (clase:$01;subclass:$06;progif:$01;name:'Serial ATA (AHCI 1.0)'),
   (clase:$01;subclass:$07;progif:$00;name:'Serial Attached SCSI'),
   (clase:$01;subclass:$80;progif:$00;name:'Other'),

   (clase:$02;subclass:$00;progif:$00;name:'Ethernet'),
   (clase:$02;subclass:$01;progif:$00;name:'Token Ring'),
   (clase:$02;subclass:$02;progif:$00;name:'FDDI'),
   (clase:$02;subclass:$03;progif:$00;name:'ATM'),
   (clase:$02;subclass:$04;progif:$00;name:'ISDN'),
   (clase:$02;subclass:$05;progif:$00;name:'WorldFip'),
   (clase:$02;subclass:$06;progif:$00;name:'PICMG 2.14'),
   (clase:$02;subclass:$80;progif:$00;name:'Other'),

   (clase:$03;subclass:$00;progif:$00;name:'VGA'),
   (clase:$03;subclass:$00;progif:$01;name:'VGA+8514'),
   (clase:$03;subclass:$01;progif:$00;name:'XGA'),
   (clase:$03;subclass:$02;progif:$00;name:'3D'),
   (clase:$03;subclass:$80;progif:$00;name:'Other'),

   (clase:$04;subclass:$00;progif:$00;name:'Video'),
   (clase:$04;subclass:$01;progif:$00;name:'Audio'),
   (clase:$04;subclass:$02;progif:$00;name:'Telephony'),
   (clase:$04;subclass:$03;progif:$00;name:'Hi-definition Audio'),
   (clase:$04;subclass:$80;progif:$00;name:'Other'),

   (clase:$05;subclass:$00;progif:$00;name:'RAM'),
   (clase:$05;subclass:$01;progif:$00;name:'Flash'),
   (clase:$05;subclass:$80;progif:$00;name:'Other'),

   (clase:$06;subclass:$00;progif:$00;name:'PCI to HOST'),
   (clase:$06;subclass:$01;progif:$00;name:'PCI to ISA'),
   (clase:$06;subclass:$02;progif:$00;name:'PCI to EISA'),
   (clase:$06;subclass:$03;progif:$00;name:'PCI to MCA'),
   (clase:$06;subclass:$04;progif:$00;name:'PCI to PCI'),
   (clase:$06;subclass:$04;progif:$01;name:'PCI to PCI (Subtractive Decode)'),
   (clase:$06;subclass:$05;progif:$00;name:'PCI to PCMCIA'),
   (clase:$06;subclass:$06;progif:$00;name:'PCI to NuBUS'),
   (clase:$06;subclass:$07;progif:$00;name:'PCI to Cardbus'),
   (clase:$06;subclass:$08;progif:$00;name:'PCI to RACEway (Transparent)'),
   (clase:$06;subclass:$08;progif:$01;name:'PCI to RACEway (End-point)'),
   (clase:$06;subclass:$09;progif:$40;name:'PCI to PCI (Primary bus to host)'),
   (clase:$06;subclass:$09;progif:$80;name:'PCI to PCI (Secondary bus to host)'),
   (clase:$06;subclass:$0A;progif:$00;name:'PCI to InfiBand'),
   (clase:$06;subclass:$80;progif:$00;name:'PCI to Other'),

   (clase:$07;subclass:$00;progif:$00;name:'Serial'),
   (clase:$07;subclass:$00;progif:$01;name:'Serial - 16450'),
   (clase:$07;subclass:$00;progif:$02;name:'Serial - 16550'),
   (clase:$07;subclass:$00;progif:$03;name:'Serial - 16650'),
   (clase:$07;subclass:$00;progif:$04;name:'Serial - 16750'),
   (clase:$07;subclass:$00;progif:$05;name:'Serial - 16850'),
   (clase:$07;subclass:$00;progif:$06;name:'Serial - 16950'),
   (clase:$07;subclass:$01;progif:$00;name:'Parallel'),
   (clase:$07;subclass:$01;progif:$01;name:'Parallel - BiDir'),
   (clase:$07;subclass:$01;progif:$02;name:'Parallel - ECP'),
   (clase:$07;subclass:$01;progif:$03;name:'Parallel - IEEE1284 Controller'),
   (clase:$07;subclass:$01;progif:$FE;name:'Parallel - IEEE1284 Target'),
   (clase:$07;subclass:$02;progif:$00;name:'Multiport Serial'),
   (clase:$07;subclass:$03;progif:$00;name:'Hayes Compatible Modem'),
   (clase:$07;subclass:$03;progif:$01;name:'Hayes Compatible Modem, 16450'),
   (clase:$07;subclass:$03;progif:$02;name:'Hayes Compatible Modem, 16550'),
   (clase:$07;subclass:$03;progif:$03;name:'Hayes Compatible Modem, 16650'),
   (clase:$07;subclass:$03;progif:$04;name:'Hayes Compatible Modem, 16750'),
   (clase:$07;subclass:$04;progif:$00;name:'GPIB (IEEE 488.1/2)'),
   (clase:$07;subclass:$05;progif:$00;name:'Smart Card'),
   (clase:$07;subclass:$80;progif:$00;name:'Other'),

{ system }
   (clase:$08;subclass:$00;progif:$00;name:'PIC'),
   (clase:$08;subclass:$00;progif:$01;name:'ISA PIC'),
   (clase:$08;subclass:$00;progif:$02;name:'EISA PIC'),
   (clase:$08;subclass:$00;progif:$10;name:'I/O APIC'),
   (clase:$08;subclass:$00;progif:$20;name:'I/O(x) APIC'),
   (clase:$08;subclass:$01;progif:$00;name:'DMA'),
   (clase:$08;subclass:$01;progif:$01;name:'ISA DMA'),
   (clase:$08;subclass:$01;progif:$02;name:'EISA DMA'),
   (clase:$08;subclass:$02;progif:$00;name:'Timer'),
   (clase:$08;subclass:$02;progif:$01;name:'ISA Timer'),
   (clase:$08;subclass:$02;progif:$02;name:'EISA Timer'),
   (clase:$08;subclass:$03;progif:$00;name:'RTC'),
   (clase:$08;subclass:$03;progif:$01;name:'ISA RTC'),
   (clase:$08;subclass:$04;progif:$00;name:'Hot-Plug'),
   (clase:$08;subclass:$05;progif:$00;name:'SD Host Controller (Without DMA)'),
   (clase:$08;subclass:$05;progif:$01;name:'SD Host Controller (With DMA)'),
   (clase:$08;subclass:$80;progif:$00;name:'Other'),

   (clase:$09;subclass:$00;progif:$00;name:'Keyboard'),
   (clase:$09;subclass:$01;progif:$00;name:'Pen'),
   (clase:$09;subclass:$02;progif:$00;name:'Mouse'),
   (clase:$09;subclass:$03;progif:$00;name:'Scanner'),
   (clase:$09;subclass:$04;progif:$00;name:'Game Port'),
   (clase:$09;subclass:$04;progif:$10;name:'Game Port (Legacy)'),
   (clase:$09;subclass:$80;progif:$00;name:'Other'),

   (clase:$0a;subclass:$00;progif:$00;name:'Generic'),
   (clase:$0a;subclass:$80;progif:$00;name:'Other'),

   (clase:$0b;subclass:$00;progif:$00;name:'386'),
   (clase:$0b;subclass:$01;progif:$00;name:'486'),
   (clase:$0b;subclass:$02;progif:$00;name:'Pentium'),
   (clase:$0b;subclass:$03;progif:$00;name:'PentiumPro'),
   (clase:$0b;subclass:$10;progif:$00;name:'DEC Alpha'),
   (clase:$0b;subclass:$20;progif:$00;name:'PowerPC'),
   (clase:$0b;subclass:$30;progif:$00;name:'MIPS'),
   (clase:$0b;subclass:$40;progif:$00;name:'Coprocessor'),
   (clase:$0b;subclass:$80;progif:$00;name:'Other'),

   (clase:$0c;subclass:$00;progif:$00;name:'FireWire'),
   (clase:$0c;subclass:$00;progif:$10;name:'OHCI FireWire'),
   (clase:$0c;subclass:$01;progif:$00;name:'ACCESS.bus'),
   (clase:$0c;subclass:$02;progif:$00;name:'SSA'),
   (clase:$0c;subclass:$03;progif:$00;name:'USB (UHCI)'),
   (clase:$0c;subclass:$03;progif:$10;name:'USB (OHCI)'),
   (clase:$0c;subclass:$03;progif:$20;name:'USB 2.0 (EHCI)'),
   (clase:$0c;subclass:$03;progif:$80;name:'USB'),
   (clase:$0c;subclass:$03;progif:$FE;name:'USB Device'),
   (clase:$0c;subclass:$04;progif:$00;name:'Fibre Channel'),
   (clase:$0c;subclass:$05;progif:$00;name:'SMBus Controller'),
   (clase:$0c;subclass:$06;progif:$00;name:'InfiniBand'),
   (clase:$0c;subclass:$07;progif:$00;name:'IPMI (SMIC)'),
   (clase:$0c;subclass:$07;progif:$01;name:'IPMI (Keyb)'),
   (clase:$0c;subclass:$07;progif:$02;name:'IPMI (Block Tx)'),
   (clase:$0c;subclass:$08;progif:$00;name:'SERCOS (IEC61491)'),
   (clase:$0c;subclass:$09;progif:$00;name:'CANbus'),
   (clase:$0c;subclass:$80;progif:$00;name:'Other'),

   (clase:$0d;subclass:$00;progif:$00;name:'iRDA Controller'),
   (clase:$0d;subclass:$01;progif:$00;name:'Consumer IR'),
   (clase:$0d;subclass:$10;progif:$00;name:'RF Controller'),
   (clase:$0d;subclass:$11;progif:$00;name:'Bluetooth'),
   (clase:$0d;subclass:$12;progif:$00;name:'Broadband'),
   (clase:$0d;subclass:$20;progif:$00;name:'Wireless Ethernet 802.11a'),
   (clase:$0d;subclass:$21;progif:$00;name:'Wireless Ethernet 802.11b'),
   (clase:$0d;subclass:$80;progif:$00;name:'Other'),

   (clase:$0e;subclass:$00;progif:$00;name:'I2O'),
   (clase:$0e;subclass:$80;progif:$00;name:'Other'),

   (clase:$0f;subclass:$01;progif:$00;name:'TV'),
   (clase:$0f;subclass:$02;progif:$00;name:'Audio'),
   (clase:$0f;subclass:$03;progif:$00;name:'Voice'),
   (clase:$0f;subclass:$04;progif:$00;name:'Data'),
   (clase:$0f;subclass:$80;progif:$00;name:'Other'),

   (clase:$10;subclass:$00;progif:$00;name:'Network'),
   (clase:$10;subclass:$10;progif:$00;name:'Entertainment'),
   (clase:$10;subclass:$80;progif:$00;name:'Other'),

   (clase:$11;subclass:$00;progif:$00;name:'DPIO Modules'),
   (clase:$11;subclass:$01;progif:$00;name:'Performance Counters'),
   (clase:$11;subclass:$10;progif:$00;name:'Comm. Sync, Time+Freq. Measurement'),
   (clase:$11;subclass:$20;progif:$00;name:'Management Card'),
   (clase:$11;subclass:$80;progif:$00;name:'Other')

   );


