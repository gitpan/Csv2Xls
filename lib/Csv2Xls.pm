package Csv2Xls;

use 5.006001;
use strict;
use warnings;
use Spreadsheet::WriteExcel;
require Exporter;

our @ISA = qw(Exporter);

our %EXPORT_TAGS = (
    'all' => [
        qw(new convert

          )
    ]
);

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } );

our @EXPORT = qw(new convert

);

our $VERSION = '0.04';

sub new() { #constructor
    my $class   = shift;
    my $hashref = shift;
    my $this    = {

        nomefile                 => $hashref->{fileName},
        allineamentotitolo       => $hashref->{titleAlign},
        allineamentostandard     => $hashref->{standardAlign},
        coloretitolo             => $hashref->{titleColor},
        larghezzacolonne         => $hashref->{columnSize},
        coloretestoprimacolonna  => $hashref->{firstColumnTextColor},
        sfondoprimariga          => $hashref->{backgroundFirstRow},
        grassettoprimacolonna    => $hashref->{boldFirstColumn},
        grassettotitolo          => $hashref->{boldTitle},
        grassettotestonormale    => $hashref->{boldStandardText},
        coloretestostandard      => $hashref->{colorStandardText},
        sfondorighenormali       => $hashref->{backgroundStandardRow},
        sfondoprimacolonna       => $hashref->{backgroundFirstColumn},
        allineamentoprimacolonna => $hashref->{alignFirstColumn}

    };

    bless $this, $class;
    return $this;
}

sub convert() { #conversion method
    my $foglio;
    my $nometab;
    my @linee;
    my @intestazione;
    my @campi;
    my $tab;
    my $titolo;
    my $standard;
    my $primacolonna;
    my $riga;
    my $colonna;
    my $campo;
    my $linea;
    my $this;
    my $csv;
    my $nomeTab;
    my $partiIntestazione;

    $this   = shift;
    $foglio = Spreadsheet::WriteExcel->new( $this->{nomefile} );

    foreach $csv (@_) {

        $nomeTab = $csv;
        $nomeTab =~ s/^.*\/|.csv//g;
        $nomeTab =~ /((^.*[_]?){1,2})(.*$)/;
        $nomeTab = $1;
        $nomeTab =~ s/_(check|report|[0-9]{6}).*$//;

        open( FILE, "$csv" ) or die "Error: $!";

        @linee        = <FILE>;
        @intestazione = split( /;/, $linee[0] );
        $tab          = $foglio->add_worksheet("$nomeTab");
        $titolo       = $foglio->add_format();
        $standard     = $foglio->add_format();
        $primacolonna = $foglio->add_format();

        if ( $this->{grassettotitolo} eq 'yes' ) {
            $titolo->set_bold();
        }

        $titolo->set_align( $this->{allineamentotitolo} );
        $titolo->set_bg_color( $this->{sfondoprimariga} );
        $titolo->set_color( $this->{coloretitolo} );
        $standard->set_color( $this->{coloretestostandard} );
        $primacolonna->set_color( $this->{coloretestoprimacolonna} );

        if ( $this->{grassettoprimacolonna} eq 'yes' ) {
            $primacolonna->set_bold();
        }

        if ( $this->{grassettotestonormale} eq 'yes' ) {
            $standard->set_bold();
        }

        $standard->set_bg_color( $this->{sfondorighenormali} );
        $primacolonna->set_bg_color( $this->{sfondoprimacolonna} );

        $primacolonna->set_align( $this->{allineamentoprimacolonna} );
        $standard->set_align( $this->{allineamentostandard} );
        $tab->set_column( 0, $#intestazione, $this->{larghezzacolonne} );
        $tab->autofilter( 0, 0, 0, $#intestazione );
        $tab->freeze_panes( 1, 1 );

        $riga    = 0;
        $colonna = 0;

        foreach $partiIntestazione (@intestazione) {

            chomp $partiIntestazione;
            $tab->write( $riga, $colonna, $partiIntestazione, $titolo );
            $colonna++;
        }
        $riga++;
        $colonna = 0;
        shift @linee;

        foreach $linea (@linee) {

            chomp $linea;
            $linea =~ s/\cM$//;
            $linea =~ s/;;/;NULL;/g;
            $linea =~ s/;;/;NULL;/g;
            $linea =~ s/;$/;NULL/;
            @campi = split( ";", $linea );

            foreach $campo (@campi) {

                chomp $campo;
                $campo =~ s/NULL//g;

                if ( $colonna == 0 ) {
                    $tab->write( $riga, $colonna, $campo, $primacolonna );
                }
                else {
                    $tab->write( $riga, $colonna, $campo, $standard );
                }
                $colonna++;
            }
            $riga++;
            $colonna = 0;
        }
    }
}

1;
__END__


=head1 NAME

Csv2Xls - Convert one or more csv to one xls

=head1 SYNOPSIS

#Example of csv1.csv file

titlecolumn1;titlecolumn2;titlecolumn3;

aaa;bbb;ccc;

fff;ggg;hhh;

-------------------------------------------------------------------

use Csv2Xls;

@fileCsv = ('csv1.csv', 'csv2.csv');

%hashRef = (

                  fileName=>'test.xls',  

                  titleAlign =>'center', 

                  standardAlign=>'left', 

                  columnSize=> 30, 

                  titleColor=> 'white', 

                  firstColumnTextColor=>'white', 

                  backgroundFirstRow=>'yellow', 

                  boldFirstColumn=>'yes', 

                  boldTitle =>'yes',  

                  boldStandardText=>'yes', 

                  colorStandardText=>'black', 

                  backgroundStandardRow=>'yellow', #or 'null' 

                  backgroundFirstColumn=>'red',  #or 'null'

                  alignFirstColumn=>'center', 

                  );

$instance1 = Csv2Xls->new(\%hashRef); #call the constructor

$instance1->convert(@fileCsv); #conversion method


=head1 AUTHOR

Cladi,  cladi@cpan.org

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2012 by Cladi Di Domenico 

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.12.3 or,
at your option, any later version of Perl 5 you may have available.


=cut
