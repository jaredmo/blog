+++
date = "2020-09-03T09:05:00-06:00"
title = "Installing Arch Linux in VMware Workstation Player"
draft = false
tags = ["Tech"]
+++

During this time when we are all isolated in our homes because of COVID-19, I have tried to make the most of it and learn some new skills. Something I've wanted to get better at is Linux. I learn best by diving into the deep end of the pool, so I decided to install [Arch Linux](https://www.archlinux.org/) in [VMware Workstation Player](https://www.vmware.com/products/workstation-player.html) as a learning experience. This blog post is to document the process for future reference. The [installation guide](https://wiki.archlinux.org/index.php/Installation_guide) on the [Arch Wiki](https://wiki.archlinux.org/) was very helpful as well, so I'd be remiss if I didn't include a link. Here goes nothing... 

First thing is to create the virtual machine in VMware. I created a 50 GB virtual hard drive, and added this next to the bottom of the *.vmx file so the machine would boot with EFI.

```
firmware = "efi"
```

Next [download](https://www.archlinux.org/download/) the latest version of Arch and boot from the ISO. You should end up at a command line. 

Verify there are results when you run the commands below. I did my install on a host machine with a wired connection. If you are performing a bare metal install on a laptop with Wi-Fi only, additional steps are required. Refer to the Arch install guide.

```
# ls /sys/firmware/efi/efivars
# ping google.com
```

If EFI is working and you have a network connection, let's move forward. Next we'll set the date and time...

```
# timedatectl set-ntp true
```

Now we'll partition the disk with *cfdisk*. Run *lsblk* to determine your particular drive path first. Mine happened to be */dev/sda*. 

```
# cfdisk /dev/sda
```

I allocated 512MB to /mnt/boot, 8GB to the swap, everything else to root. Next we'll format the drive.

```
# mkfs.vfat -F32 /dev/sda1
# mkswap /dev/sda2
# swapon /dev/sda2
# mkfs.ext4 /dev/sda3
```

Now we'll create the boot and home directories, as well as mount the partitions...

```
# mkdir /mnt/boot
# mkdir /mnt/home
# mount /dev/sda1 /mnt/boot
# mount /dev/sda3 /mnt
```

Before going forward let's check our work by running *lsblk*. Once you are happy with the drive move forward.

Now we'll recreate the *mirrorlist* file based on download speed. This will rate and sort the 10 most recently synchronized mirrors by download speed, then overwrite */etc/pacman.d/mirrorlist*.

```
# reflector --verbose --latest 10 --sort rate --save /etc/pacman.d/mirrorlist
```

Now the moment of truth. Let's install Arch. 

```
# pacstrap /mnt base linux linux-firmware
```

Next we'll generate and examine the file system.

```
# genfstab /mnt >> /mnt/etc/fstab
# nano /mnt/etc/fstab
```

Arch is now installed on the hard drive. Let's login to the new system, and install some packages. I've included a couple text editors, networking, sudo, git, grub bootloader, and the linux-lts kernel. Using this kernel will (hopefully) increase system stability down the road.

```
# arch-chroot /mnt
# pacman -S dhcpcd efibootmgr git grub linux-lts nano sudo vi
```

Set the time zone. I happen to be on US Central time.

```
# ln -sf /usr/share/zoneinfo/America/Chicago /etc/localtime
# hwclock --systohc
```

Now we'll generate the locale, and uncomment *en_US.UTF-8 UTF-8* in *locale.gen*, and configure the hostname. Mine is *archVM*.

```
# locale-gen
# nano /etc/locale.gen # Uncomment en_US.UTF-8 UTF-8
# echo "LANG=en_US.UTF-8" >> /etc/locale.conf
# echo "archVM" >> /etc/hostname
```

Next is setting up the hosts file.

```
# nano /etc/hosts
```

Add the following text and save. 

```
127.0.0.1	localhost
::1		localhost
127.0.1.1	archVM.localdomain	archVM
```

Now we'll set the root account password.

```
# passwd
```

Install the grub bootloader.

```
# grub-install --target=x86_64-efi --efi-directory=/boot --bootloader-id=GRUB
# grub-mkconfig -o /boot/grub/grub.cfg
```

...And enable networking

```
# systemctl enable dhcpcd.service
```

After all that's done, let's unmount the hard disk, reboot, and cross our fingers. Be sure to unmount the installation ISO at the point.

```
# exit
# umount -R /mnt
# reboot
```

Once the system reboots, you should have a login prompt. Congrats! You've installed Arch. Everything else from this point forward is icing on the cake.

I happen to like the [KDE Plasma](https://kde.org/plasma-desktop) desktop environment and [yay](https://github.com/Jguer/yay) AUR helper. The next steps can be used to get those running, as well as installing VMware's guest tools.

First let's update Arch and create a non-root user account.

```
# pacman -Syu
# useradd -m [username]
# passwd [username]
# visudo #Uncomment relevant lines to allow all users to run sudo. Save and quit.
```

Now we'll install xorg and our desktop environment, as well as enable the login screen and network manager. This will take awhile depending on your network speed...

```
# pacman -S xorg
# pacman -S plasma-meta kde-applications
# systemctl enable sddm.service
# systemctl enable NetworkManager.service
```

Now let's install VMware guest tools as well as some drivers that will help us out. After that finishes issue the *reboot*.

```
# pacman -S open-vm-tools
# pacman -S xf86-input-vmmouse xf86-video-vmware mesa gtk2 gtkmm
# systemctl enable vmtoolsd.service
# systemctl enable vmware-vmblock-fuse.service
# reboot
```

You should now see a login screen for KDE Plasma. It's starting the feel more like a finished system. Let's login with as the newly created non-root user.

Open up Konsole or whatever terminal you like, now let's install yay. If you don't intend to use the [AUR](https://aur.archlinux.org/), you can stop here.

```
$ cd /opt
$ git clone https://aur.archlinux.org/yay.git
$ sudo chown -R [username]:[username] yay
$ cd yay
$ makepkg -si
```

Now you can update your entire system just by opening a terminal and issuing the *yay* command. That feels pretty good! 

I hope this little guide was helpful. I certainly learned a lot in the process, and am glad to have completed the exercise.
