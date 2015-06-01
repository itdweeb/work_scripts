#!/usr/bin/perl
# Simple DHCP client - sending a broadcasted DHCP Discover request

use IO::Socket::INET;
use Net::DHCP::Packet;
use Net::DHCP::Constants;

# create DHCP Packet
$discover = Net::DHCP::Packet->new(
	xid => int(rand(0x12345678)), # random xid
	Flags => 0x8000,              # ask for broadcast
	DHO_DHCP_MESSAGE_TYPE() => DHCPDISCOVER()
	);

# send packet
$handle = IO::Socket::INET->new(Proto => 'udp',
	Broadcast => 1,
	PeerPort  => '67',
	LocalPort => '68',
	LocalAddr => 'chris-pc',
	PeerAddr  => '255.255.255.255')
	or die "socket: $@";     # yes, it uses $@ here
$handle->send($discover->serialize())
	or die "Error sending broadcast inform:$!\n";
