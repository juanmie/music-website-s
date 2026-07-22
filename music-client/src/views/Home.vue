<template>
  <div class="theme-toggle" @click="toggleTheme" >
    <el-icon :size="20"><Moon v-if="!isDark" /><Sunny v-else /></el-icon>
  </div>

  <!--轮播图-->
  <el-carousel v-if="swiperList.length" class="swiper-container" type="card" height="20vw" :interval="4000">
    <el-carousel-item v-for="(item, index) in swiperList" :key="index">
      <img :src="HttpManager.attachImageUrl(item.pic)" alt=""/>
    </el-carousel-item>
  </el-carousel>
  <!--热门歌单-->
  <play-list class="play-list-container" title="歌单" path="song-sheet-detail" :playList="songList"></play-list>
  <!--热门歌手-->
  <play-list class="play-list-container" title="歌手" path="singer-detail" :playList="singerList"></play-list>
</template>

<script lang="ts" setup>
import {ref, onMounted, nextTick} from "vue";

import PlayList from "@/components/PlayList.vue";
import {NavName} from "@/enums";
import {HttpManager} from "@/api";
import mixin from "@/mixins/mixin";
import { useDark, useToggle } from '@vueuse/core';
import { Moon, Sunny } from '@element-plus/icons-vue';
const isDark = useDark();
const toggleTheme = (event: MouseEvent) => {
// 记录切换前的主题状态
  const wasDark = isDark.value;

  if (!document.startViewTransition) {
    isDark.value = !wasDark;
    // 可选添加渐隐渐现效果
    document.documentElement.style.transition = 'background 0.3s';
    return;
  }
// 保存点击位置坐标
  const { clientX, clientY } = event;

  // 计算覆盖全屏所需半径（勾股定理计算对角线距离）
  const radius = Math.hypot(
      Math.max(clientX, innerWidth - clientX),
      Math.max(clientY, innerHeight - clientY)
  );
  // 启动视图过渡
  const transition = document.startViewTransition(() => {
    isDark.value = !wasDark;  // 使用切换前的状态确保值正确
  });


  transition.ready.then(() => {
    // 设置动画起始点坐标
    const root = document.documentElement;
    root.style.setProperty('--x', `${clientX}px`);
    root.style.setProperty('--y', `${clientY}px`);

    // 动态设置动画方向
    const [startRadius, endRadius] = !wasDark
        ? [radius, 0]  // 暗→亮：从全屏收缩到点击点‌:ml-citation{ref="3" data="citationList"}
        : [0, radius] ;   // 亮→暗：从点击点扩散到全屏‌:ml-citation{ref="1" data="citationList"}


// 确定作用的目标伪元素
    const pseudoElement = wasDark
        ? '::view-transition-new(root)' // 作用于旧视图（暗色背景）
        : '::view-transition-new(root)';// 作用于新视图（暗色背景）

//执行动画
    root.animate(
        {
          clipPath: [
            `circle(${startRadius}px at var(--x) var(--y))`,
            `circle(${endRadius}px at var(--x) var(--y))`
          ]
        },
        {
          duration: 1500,
          easing: 'cubic-bezier(0.4, 0, 0.2, 1)',
          pseudoElement
        }
    );
  });
};



const songList = ref([]); // 歌单列表
const singerList = ref([]); // 歌手列表
const swiperList = ref([]);// 轮播图 每次都在进行查询
const {changeIndex} = mixin();
try {

  HttpManager.getBannerList().then((res) => {
    swiperList.value = (res as ResponseBody).data.sort();
  });

  HttpManager.getSongList().then((res) => {
    songList.value = (res as ResponseBody).data.sort().slice(0, 10);
  });

  HttpManager.getAllSinger().then((res) => {
    singerList.value = (res as ResponseBody).data.sort().slice(0, 10);
  });

  onMounted(() => {
    changeIndex(NavName.Home);
  });
} catch (error) {
  console.error(error);
}
</script>




<style lang="scss" scoped>
@import "@/assets/css/var.scss";

/*轮播图*/
.swiper-container {
  width: 90%;
  margin: auto;
  padding-top: 20px;

  img {
    width: 100%;
  }
}

.swiper-container:deep(.el-carousel__indicators.el-carousel__indicators--outside) {
  display: inline-block;
  transform: translateX(30vw);
}

.el-slider__runway {
  background-color: $color-blue;
}


/* 强制覆盖默认动画 */
::view-transition-old(root),
::view-transition-new(root) {
  animation: none;
  mix-blend-mode: normal;

}

::view-transition-old(root) {
  z-index:9999;  /* 收缩时旧主题保持在顶层‌*/
}


.theme-toggle {
  @apply fixed bottom-10 left-10 p-3 rounded-full bg-white dark:bg-gray-800 shadow-lg cursor-pointer hover:scale-110 transition-transform;
}




</style>