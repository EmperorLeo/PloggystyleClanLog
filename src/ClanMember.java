import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

public class ClanMember {

	private String characterName;
	private ItemInformation itemInfo;
	private boolean active;
	
	public ClanMember(String characterName) {
		this.characterName = characterName;
	}
	
	public ItemInformation getInformation() {		
		ItemInformation info = new ItemInformation();
		
		try {
			System.out.println("Getting information for: " + characterName);
			Document mainPage = Jsoup.connect("http://na-bns.ncsoft.com/ingame/bs/character/profile?c=" + URLEncoder.encode(characterName, "UTF-8")).get();
			if(!mainPage.select(".desc .guild").text().equals("Ploggystyle")) {
				return null;
			}
			active = true;
			info.setMasteryLevel(mainPage.select(".masteryLv").text().split(" ")[1]);
			
			Document itemPage = Jsoup.connect("http://na-bns.ncsoft.com/ingame/bs/character/data/equipments?c=" + URLEncoder.encode(characterName, "UTF-8")).get();
			info.setWeapon(itemPage.select(".wrapWeapon .name").text());
			info.setSoul(itemPage.select(".wrapAccessory.soul .name").text());
			info.setRing(itemPage.select(".wrapAccessory.ring .name").text());
			info.setNecklace(itemPage.select(".wrapAccessory.necklace .name").text());
			info.setEarring(itemPage.select(".wrapAccessory.earring .name").text());
			info.setBelt(itemPage.select(".wrapAccessory.belt .name").text());
			info.setBracelet(itemPage.select(".wrapAccessory.bracelet .name").text());
			info.setPet(itemPage.select(".wrapAccessory.guard .name").text());
			info.setSoulBadge(itemPage.select(".wrapAccessory.singongpae .name").text());
			info.setMysticBadge(itemPage.select(".wrapAccessory.rune .name").text());
			
			itemInfo = info;			
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return getItemInformation();
	}
	
	public ItemInformation getItemInformation() {
		return itemInfo;
	}
	
	public boolean getActive() {
		return active;
	}
	
	public String getCharacterName() {
		return characterName;
	}
}
